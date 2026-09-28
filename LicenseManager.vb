Imports System
Imports System.Globalization
Imports System.IO
Imports System.Security.Cryptography
Imports System.Security.Principal
Imports System.Text
Imports Microsoft.Win32

Namespace CreoRegenAppVB

    Friend Enum LicenseState
        Trial = 0
        Licensed = 1
        Expired = 2
    End Enum

    Friend NotInheritable Class LicenseStatus
        Friend Property State As LicenseState
        Friend Property DaysRemaining As Integer
        Friend Property Licensee As String
    End Class

    ''' <summary>
    ''' Resolves the current entitlement: an installed license key, or the 30-day trial.
    ''' </summary>
    ''' <remarks>
    ''' The trial HMAC secret ships inside the binary, so trial tamper-detection is a
    ''' deterrent only. Key forgery is the part that is genuinely prevented: the ECDSA
    ''' private key never leaves the key generator.
    ''' </remarks>
    Friend NotInheritable Class LicenseManager

        Friend Const TrialLengthDays As Integer = 30

        Private Const AppDataFolderName As String = "BOMForge"
        Private Const LicenseFileName As String = "license.lic"
        Private Const StateFileName As String = "state.dat"
        Private Const RegistryKeyPath As String = "Software\BOMForge\Settings"
        Private Const FirstRunValueName As String = "UiState"
        Private Const LastSeenValueName As String = "LastState"
        Private Const SignatureValueName As String = "Sig"
        Private Const TrialSecretSalt As String = "BOMForge.Trial.v1"

        ' Absorbs small clock corrections without burning a legitimate trial.
        Private Shared ReadOnly ClockTolerance As TimeSpan = TimeSpan.FromHours(24)

        Private Sub New()
        End Sub

        Friend Shared ReadOnly Property AppDataFolder As String
            Get
                Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), AppDataFolderName)
            End Get
        End Property

        Friend Shared ReadOnly Property LicenseFilePath As String
            Get
                Return Path.Combine(AppDataFolder, LicenseFileName)
            End Get
        End Property

        Friend Shared Function Evaluate() As LicenseStatus
            Dim installed = ReadInstalledLicense()
            If installed IsNot Nothing Then
                Return New LicenseStatus With {
                    .State = LicenseState.Licensed,
                    .DaysRemaining = 0,
                    .Licensee = installed.Licensee
                }
            End If

            Return EvaluateTrial()
        End Function

        Friend Shared Function TryActivate(keyText As String, ByRef info As LicenseInfo) As Boolean
            info = Nothing

            Dim parsed As LicenseInfo = Nothing
            If Not LicenseFormat.TryParse(keyText, parsed) Then
                Return False
            End If

            Try
                Directory.CreateDirectory(AppDataFolder)
                File.WriteAllText(LicenseFilePath, keyText.Trim(), New UTF8Encoding(False))
            Catch ex As Exception
                CreoSessionService.LogApplicationError("Save license", ex)
                Return False
            End Try

            info = parsed
            Return True
        End Function

        Friend Shared Sub RecordHeartbeat()
            Try
                If ReadInstalledLicense() IsNot Nothing Then
                    Return
                End If

                EvaluateTrial()
            Catch ex As Exception
                CreoSessionService.LogApplicationError("Trial heartbeat", ex)
            End Try
        End Sub

        Private Shared Function ReadInstalledLicense() As LicenseInfo
            Try
                Dim licensePath = LicenseFilePath
                If Not File.Exists(licensePath) Then
                    Return Nothing
                End If

                Dim keyText = File.ReadAllText(licensePath)
                Dim parsed As LicenseInfo = Nothing
                If LicenseFormat.TryParse(keyText, parsed) Then
                    Return parsed
                End If
            Catch ex As Exception
                CreoSessionService.LogApplicationError("Read license", ex)
            End Try

            Return Nothing
        End Function

        Private Shared Function EvaluateTrial() As LicenseStatus
            Dim nowUtc = Date.UtcNow

            Dim registryRecord As TrialRecord = Nothing
            Dim registryCorrupt = Not TryReadRegistryRecord(registryRecord)

            Dim fileRecord As TrialRecord = Nothing
            Dim fileCorrupt = Not TryReadFileRecord(fileRecord)

            Dim tampered = registryCorrupt OrElse fileCorrupt

            If registryRecord Is Nothing AndAlso fileRecord Is Nothing Then
                Dim fresh As New TrialRecord With {.FirstRunUtc = nowUtc, .LastSeenUtc = nowUtc}
                WriteRecord(fresh)
                Return BuildTrialStatus(fresh, nowUtc)
            End If

            ' Earliest start and latest sighting win, so removing one mirror cannot reset the trial.
            Dim firstRun = EarliestOf(registryRecord, fileRecord)
            Dim lastSeen = LatestOf(registryRecord, fileRecord)

            If nowUtc < lastSeen.Subtract(ClockTolerance) Then
                tampered = True
            End If

            If firstRun > nowUtc.Add(ClockTolerance) Then
                tampered = True
            End If

            ' High-water mark: winding the clock forward burns the trial permanently.
            Dim reference = If(lastSeen > nowUtc, lastSeen, nowUtc)

            If tampered Then
                firstRun = reference.AddDays(-TrialLengthDays)
            End If

            Dim current As New TrialRecord With {.FirstRunUtc = firstRun, .LastSeenUtc = reference}
            WriteRecord(current)
            Return BuildTrialStatus(current, reference)
        End Function

        Private Shared Function BuildTrialStatus(record As TrialRecord, reference As Date) As LicenseStatus
            Dim elapsedDays = (reference - record.FirstRunUtc).TotalDays
            Dim remaining = CInt(Math.Ceiling(TrialLengthDays - elapsedDays))

            If remaining <= 0 Then
                Return New LicenseStatus With {
                    .State = LicenseState.Expired,
                    .DaysRemaining = 0,
                    .Licensee = Nothing
                }
            End If

            If remaining > TrialLengthDays Then
                remaining = TrialLengthDays
            End If

            Return New LicenseStatus With {
                .State = LicenseState.Trial,
                .DaysRemaining = remaining,
                .Licensee = Nothing
            }
        End Function

        Private Shared Function EarliestOf(first As TrialRecord, second As TrialRecord) As Date
            If first Is Nothing Then Return second.FirstRunUtc
            If second Is Nothing Then Return first.FirstRunUtc
            Return If(first.FirstRunUtc <= second.FirstRunUtc, first.FirstRunUtc, second.FirstRunUtc)
        End Function

        Private Shared Function LatestOf(first As TrialRecord, second As TrialRecord) As Date
            If first Is Nothing Then Return second.LastSeenUtc
            If second Is Nothing Then Return first.LastSeenUtc
            Return If(first.LastSeenUtc >= second.LastSeenUtc, first.LastSeenUtc, second.LastSeenUtc)
        End Function

        Private Shared Sub WriteRecord(record As TrialRecord)
            Dim signature = ComputeSignature(record.FirstRunUtc.Ticks, record.LastSeenUtc.Ticks)

            Try
                Using key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath)
                    If key IsNot Nothing Then
                        key.SetValue(FirstRunValueName, record.FirstRunUtc.Ticks.ToString(CultureInfo.InvariantCulture), RegistryValueKind.String)
                        key.SetValue(LastSeenValueName, record.LastSeenUtc.Ticks.ToString(CultureInfo.InvariantCulture), RegistryValueKind.String)
                        key.SetValue(SignatureValueName, signature, RegistryValueKind.String)
                    End If
                End Using
            Catch ex As Exception
                CreoSessionService.LogApplicationError("Write trial registry state", ex)
            End Try

            Try
                Directory.CreateDirectory(AppDataFolder)
                Dim line = String.Join("|", record.FirstRunUtc.Ticks.ToString(CultureInfo.InvariantCulture),
                                            record.LastSeenUtc.Ticks.ToString(CultureInfo.InvariantCulture),
                                            signature)
                File.WriteAllText(Path.Combine(AppDataFolder, StateFileName), line, New UTF8Encoding(False))
            Catch ex As Exception
                CreoSessionService.LogApplicationError("Write trial file state", ex)
            End Try
        End Sub

        ''' <summary>Returns False only when a record is present but fails validation.</summary>
        Private Shared Function TryReadRegistryRecord(ByRef record As TrialRecord) As Boolean
            record = Nothing

            Try
                Using key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, False)
                    If key Is Nothing Then
                        Return True
                    End If

                    Dim firstRunText = TryCast(key.GetValue(FirstRunValueName), String)
                    Dim lastSeenText = TryCast(key.GetValue(LastSeenValueName), String)
                    Dim signature = TryCast(key.GetValue(SignatureValueName), String)

                    If firstRunText Is Nothing AndAlso lastSeenText Is Nothing AndAlso signature Is Nothing Then
                        Return True
                    End If

                    Return TryBuildRecord(firstRunText, lastSeenText, signature, record)
                End Using
            Catch ex As Exception
                CreoSessionService.LogApplicationError("Read trial registry state", ex)
                Return True
            End Try
        End Function

        ''' <summary>Returns False only when a record is present but fails validation.</summary>
        Private Shared Function TryReadFileRecord(ByRef record As TrialRecord) As Boolean
            record = Nothing

            Dim statePath = Path.Combine(AppDataFolder, StateFileName)
            Dim content As String
            Try
                If Not File.Exists(statePath) Then
                    Return True
                End If

                content = File.ReadAllText(statePath)
            Catch ex As Exception
                CreoSessionService.LogApplicationError("Read trial file state", ex)
                Return True
            End Try

            Dim parts = content.Trim().Split("|"c)
            If parts.Length <> 3 Then
                Return False
            End If

            Return TryBuildRecord(parts(0), parts(1), parts(2), record)
        End Function

        Private Shared Function TryBuildRecord(firstRunText As String, lastSeenText As String, signature As String, ByRef record As TrialRecord) As Boolean
            record = Nothing

            Dim firstRunTicks As Long
            Dim lastSeenTicks As Long
            If Not Long.TryParse(firstRunText, NumberStyles.None, CultureInfo.InvariantCulture, firstRunTicks) Then
                Return False
            End If

            If Not Long.TryParse(lastSeenText, NumberStyles.None, CultureInfo.InvariantCulture, lastSeenTicks) Then
                Return False
            End If

            If firstRunTicks < 0 OrElse firstRunTicks > Date.MaxValue.Ticks Then
                Return False
            End If

            If lastSeenTicks < 0 OrElse lastSeenTicks > Date.MaxValue.Ticks Then
                Return False
            End If

            If Not SignaturesMatch(signature, ComputeSignature(firstRunTicks, lastSeenTicks)) Then
                Return False
            End If

            record = New TrialRecord With {
                .FirstRunUtc = New Date(firstRunTicks, DateTimeKind.Utc),
                .LastSeenUtc = New Date(lastSeenTicks, DateTimeKind.Utc)
            }
            Return True
        End Function

        Private Shared Function SignaturesMatch(left As String, right As String) As Boolean
            If left Is Nothing OrElse right Is Nothing OrElse left.Length <> right.Length Then
                Return False
            End If

            Dim difference As Integer = 0
            For index = 0 To left.Length - 1
                difference = difference Or (AscW(left(index)) Xor AscW(right(index)))
            Next

            Return difference = 0
        End Function

        Private Shared Function ComputeSignature(firstRunTicks As Long, lastSeenTicks As Long) As String
            Dim identity = GetUserIdentity()
            Dim message = String.Join("|", firstRunTicks.ToString(CultureInfo.InvariantCulture),
                                           lastSeenTicks.ToString(CultureInfo.InvariantCulture),
                                           identity)

            Using keyHash As SHA256 = SHA256.Create()
                Dim hmacKey = keyHash.ComputeHash(Encoding.UTF8.GetBytes(TrialSecretSalt & "|" & identity))
                Using hmac As New HMACSHA256(hmacKey)
                    Return Convert.ToBase64String(hmac.ComputeHash(Encoding.UTF8.GetBytes(message)))
                End Using
            End Using
        End Function

        Private Shared Function GetUserIdentity() As String
            Try
                Using current = WindowsIdentity.GetCurrent()
                    If current IsNot Nothing AndAlso current.User IsNot Nothing Then
                        Return current.User.Value
                    End If
                End Using
            Catch
            End Try

            Return Environment.UserName & "@" & Environment.MachineName
        End Function

        Private NotInheritable Class TrialRecord
            Friend Property FirstRunUtc As Date
            Friend Property LastSeenUtc As Date
        End Class
    End Class
End Namespace
