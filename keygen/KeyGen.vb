Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Forms

Namespace CreoRegenAppVB

    Friend Module KeyGenProgram
        <STAThread>
        Public Sub Main()
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            Application.Run(New KeyGenForm())
        End Sub
    End Module

    ''' <summary>
    ''' Offline issuer for BOMForge license keys. The ECDSA private key it creates never
    ''' leaves this machine and must never be shipped with the application.
    ''' </summary>
    Friend Class KeyGenForm
        Inherits Form

        Private Const StoreFolderName As String = "BOMForge.KeyGen"
        Private Const PrivateKeyFileName As String = "private.blob"
        Private Const IssuedLogFileName As String = "issued-keys.csv"

        Private ReadOnly _publicKeyTextBox As TextBox
        Private ReadOnly _licenseeTextBox As TextBox
        Private ReadOnly _generatedKeyTextBox As TextBox
        Private ReadOnly _statusLabel As Label
        Private ReadOnly _generateButton As Button
        Private ReadOnly _saveButton As Button

        Private _privateBlob As Byte()

        Friend Sub New()
            Text = "BOMForge Key Generator"
            Width = 680
            Height = 560
            MinimumSize = New Size(560, 460)
            StartPosition = FormStartPosition.CenterScreen

            Dim root As New TableLayoutPanel()
            root.Dock = DockStyle.Fill
            root.Padding = New Padding(12)
            root.ColumnCount = 1
            root.RowCount = 7
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 40.0F))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 60.0F))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            Dim publicKeyLabel As New Label()
            publicKeyLabel.AutoSize = True
            publicKeyLabel.Text = "Public key — paste this into LicenseFormat.PublicKeyBase64 and rebuild BOMForge:"
            publicKeyLabel.Margin = New Padding(0, 0, 0, 4)

            _publicKeyTextBox = New TextBox()
            _publicKeyTextBox.Multiline = True
            _publicKeyTextBox.ReadOnly = True
            _publicKeyTextBox.ScrollBars = ScrollBars.Vertical
            _publicKeyTextBox.Dock = DockStyle.Fill
            _publicKeyTextBox.Font = New Font(FontFamily.GenericMonospace, 9.0F)

            Dim publicKeyButtons As New FlowLayoutPanel()
            publicKeyButtons.Dock = DockStyle.Fill
            publicKeyButtons.AutoSize = True
            publicKeyButtons.AutoSizeMode = AutoSizeMode.GrowAndShrink
            publicKeyButtons.Margin = New Padding(0, 4, 0, 10)

            Dim copyPublicKeyButton As New Button()
            copyPublicKeyButton.Text = "Copy Public Key"
            copyPublicKeyButton.AutoSize = True
            AddHandler copyPublicKeyButton.Click, Sub(sender, args) CopyToClipboard(_publicKeyTextBox.Text)
            publicKeyButtons.Controls.Add(copyPublicKeyButton)

            Dim licenseeLabel As New Label()
            licenseeLabel.AutoSize = True
            licenseeLabel.Text = "Licensee name (max " & LicenseFormat.MaxLicenseeLength.ToString() & " characters):"

            Dim issueRow As New TableLayoutPanel()
            issueRow.Dock = DockStyle.Fill
            issueRow.AutoSize = True
            issueRow.AutoSizeMode = AutoSizeMode.GrowAndShrink
            issueRow.ColumnCount = 2
            issueRow.RowCount = 1
            issueRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            issueRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            issueRow.Margin = New Padding(0, 2, 0, 10)

            _licenseeTextBox = New TextBox()
            _licenseeTextBox.Dock = DockStyle.Fill
            _licenseeTextBox.MaxLength = LicenseFormat.MaxLicenseeLength

            _generateButton = New Button()
            _generateButton.Text = "Generate Key"
            _generateButton.AutoSize = True
            _generateButton.Margin = New Padding(6, 0, 0, 0)
            AddHandler _generateButton.Click, AddressOf GenerateClicked

            issueRow.Controls.Add(_licenseeTextBox, 0, 0)
            issueRow.Controls.Add(_generateButton, 1, 0)

            _generatedKeyTextBox = New TextBox()
            _generatedKeyTextBox.Multiline = True
            _generatedKeyTextBox.ReadOnly = True
            _generatedKeyTextBox.ScrollBars = ScrollBars.Vertical
            _generatedKeyTextBox.Dock = DockStyle.Fill
            _generatedKeyTextBox.Font = New Font(FontFamily.GenericMonospace, 9.0F)

            Dim resultButtons As New FlowLayoutPanel()
            resultButtons.Dock = DockStyle.Fill
            resultButtons.AutoSize = True
            resultButtons.AutoSizeMode = AutoSizeMode.GrowAndShrink
            resultButtons.Margin = New Padding(0, 4, 0, 0)

            Dim copyKeyButton As New Button()
            copyKeyButton.Text = "Copy Key"
            copyKeyButton.AutoSize = True
            AddHandler copyKeyButton.Click, Sub(sender, args) CopyToClipboard(_generatedKeyTextBox.Text)

            _saveButton = New Button()
            _saveButton.Text = "Save .lic File..."
            _saveButton.AutoSize = True
            _saveButton.Enabled = False
            AddHandler _saveButton.Click, AddressOf SaveClicked

            _statusLabel = New Label()
            _statusLabel.AutoSize = True
            _statusLabel.Margin = New Padding(12, 6, 0, 0)

            resultButtons.Controls.Add(copyKeyButton)
            resultButtons.Controls.Add(_saveButton)
            resultButtons.Controls.Add(_statusLabel)

            root.Controls.Add(publicKeyLabel, 0, 0)
            root.Controls.Add(_publicKeyTextBox, 0, 1)
            root.Controls.Add(publicKeyButtons, 0, 2)
            root.Controls.Add(licenseeLabel, 0, 3)
            root.Controls.Add(issueRow, 0, 4)
            root.Controls.Add(_generatedKeyTextBox, 0, 5)
            root.Controls.Add(resultButtons, 0, 6)

            Controls.Add(root)
        End Sub

        Private Shared ReadOnly Property StoreFolder As String
            Get
                Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), StoreFolderName)
            End Get
        End Property

        Private Shared ReadOnly Property PrivateKeyPath As String
            Get
                Return Path.Combine(StoreFolder, PrivateKeyFileName)
            End Get
        End Property

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            LoadOrCreateSigningKey()
        End Sub

        Private Sub LoadOrCreateSigningKey()
            If File.Exists(PrivateKeyPath) Then
                Try
                    Dim protectedBlob = Convert.FromBase64String(File.ReadAllText(PrivateKeyPath))
                    _privateBlob = ProtectedData.Unprotect(protectedBlob, Nothing, DataProtectionScope.CurrentUser)
                    _publicKeyTextBox.Text = ExportPublicKeyBase64(_privateBlob)
                    SetStatus("Signing key loaded.", False)
                Catch ex As Exception
                    _generateButton.Enabled = False
                    SetStatus("Signing key could not be read.", True)
                    MessageBox.Show(Me,
                        "The signing key at " & PrivateKeyPath & " exists but could not be decrypted." & Environment.NewLine & Environment.NewLine &
                        "It is bound to the Windows user that created it. Restore it from backup — regenerating a new key pair would invalidate every license key already issued." & Environment.NewLine & Environment.NewLine &
                        ex.Message,
                        "Signing Key Unavailable", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                Return
            End If

            Dim answer = MessageBox.Show(Me,
                "No signing key was found." & Environment.NewLine & Environment.NewLine &
                "Generate a new ECDSA P-256 key pair now?" & Environment.NewLine & Environment.NewLine &
                "Do this only once. Back up " & PrivateKeyPath & " afterwards — if it is lost, no further keys can be issued for the current build of BOMForge.",
                "Generate Signing Key", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If answer <> DialogResult.Yes Then
                _generateButton.Enabled = False
                SetStatus("No signing key — cannot issue license keys.", True)
                Return
            End If

            Try
                Dim creationParameters As New CngKeyCreationParameters()
                creationParameters.ExportPolicy = CngExportPolicies.AllowPlaintextExport
                creationParameters.KeyUsage = CngKeyUsages.Signing

                Using created = CngKey.Create(CngAlgorithm.ECDsaP256, Nothing, creationParameters)
                    _privateBlob = created.Export(CngKeyBlobFormat.EccPrivateBlob)
                    _publicKeyTextBox.Text = Convert.ToBase64String(created.Export(CngKeyBlobFormat.EccPublicBlob))
                End Using

                Directory.CreateDirectory(StoreFolder)
                Dim protectedBlob = ProtectedData.Protect(_privateBlob, Nothing, DataProtectionScope.CurrentUser)
                File.WriteAllText(PrivateKeyPath, Convert.ToBase64String(protectedBlob), New UTF8Encoding(False))

                SetStatus("Signing key created.", False)
                MessageBox.Show(Me,
                    "Key pair created and saved to " & PrivateKeyPath & "." & Environment.NewLine & Environment.NewLine &
                    "Copy the public key above into LicenseFormat.PublicKeyBase64 and rebuild BOMForge before issuing keys.",
                    "Signing Key Created", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                _privateBlob = Nothing
                _generateButton.Enabled = False
                SetStatus("Signing key could not be created.", True)
                MessageBox.Show(Me, ex.Message, "Signing Key Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Shared Function ExportPublicKeyBase64(privateBlob As Byte()) As String
            Using key = CngKey.Import(privateBlob, CngKeyBlobFormat.EccPrivateBlob)
                Return Convert.ToBase64String(key.Export(CngKeyBlobFormat.EccPublicBlob))
            End Using
        End Function

        Private Sub GenerateClicked(sender As Object, e As EventArgs)
            _generatedKeyTextBox.Clear()
            _saveButton.Enabled = False

            If _privateBlob Is Nothing Then
                SetStatus("No signing key available.", True)
                Return
            End If

            Dim licensee = LicenseFormat.NormalizeLicensee(_licenseeTextBox.Text)
            If licensee Is Nothing Then
                SetStatus("Enter a licensee name (no control characters).", True)
                Return
            End If

            Try
                Dim licenseId(LicenseFormat.LicenseIdLength - 1) As Byte
                Using rng As New RNGCryptoServiceProvider()
                    rng.GetBytes(licenseId)
                End Using

                Dim payload = LicenseFormat.BuildPayload(licenseId, licensee, Date.UtcNow)

                Dim signature As Byte()
                Using key = CngKey.Import(_privateBlob, CngKeyBlobFormat.EccPrivateBlob)
                    Using signer As New ECDsaCng(key)
                        signer.HashAlgorithm = CngAlgorithm.Sha256
                        signature = signer.SignData(payload)
                    End Using
                End Using

                Dim keyText = LicenseFormat.ComposeKey(payload, signature)

                Dim verified As LicenseInfo = Nothing
                If Not LicenseFormat.TryParse(keyText, verified) Then
                    SetStatus("Key generated but failed self-verification — is PublicKeyBase64 up to date?", True)
                Else
                    SetStatus("Key issued for " & verified.Licensee & ".", False)
                End If

                _generatedKeyTextBox.Text = keyText
                _saveButton.Enabled = True

                Dim licenseIdText As New StringBuilder(licenseId.Length * 2)
                For Each value In licenseId
                    licenseIdText.Append(value.ToString("X2"))
                Next

                AppendIssuedLog(licenseIdText.ToString(), licensee, keyText)
            Catch ex As Exception
                SetStatus("Key generation failed.", True)
                MessageBox.Show(Me, ex.Message, "Key Generation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub SaveClicked(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(_generatedKeyTextBox.Text) Then
                Return
            End If

            Using dialog As New SaveFileDialog()
                dialog.Filter = "License files (*.lic)|*.lic"
                dialog.Title = "Save License File"
                dialog.FileName = "bomforge.lic"
                dialog.DefaultExt = "lic"
                dialog.AddExtension = True
                dialog.OverwritePrompt = True

                If dialog.ShowDialog(Me) <> DialogResult.OK Then
                    Return
                End If

                Try
                    File.WriteAllText(dialog.FileName, _generatedKeyTextBox.Text, New UTF8Encoding(False))
                    SetStatus("License file saved.", False)
                Catch ex As Exception
                    SetStatus("License file could not be saved.", True)
                    MessageBox.Show(Me, ex.Message, "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Sub

        Private Sub AppendIssuedLog(licenseId As String, licensee As String, keyText As String)
            Try
                Directory.CreateDirectory(StoreFolder)
                Dim logPath = Path.Combine(StoreFolder, IssuedLogFileName)
                Dim writeHeader = Not File.Exists(logPath)

                Using writer As New StreamWriter(logPath, True, New UTF8Encoding(False))
                    If writeHeader Then
                        writer.WriteLine(String.Join(",", EscapeCsvValues(New String() {"IssuedUtc", "LicenseId", "Licensee", "Key"})))
                    End If

                    Dim values As String() = {
                        Date.UtcNow.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                        licenseId,
                        licensee,
                        keyText
                    }
                    writer.WriteLine(String.Join(",", EscapeCsvValues(values)))
                End Using
            Catch ex As Exception
                SetStatus("Key issued, but the issue log could not be written.", True)
                MessageBox.Show(Me, ex.Message, "Log Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        Private Shared Function EscapeCsvValues(values As IEnumerable(Of String)) As IEnumerable(Of String)
            Dim escaped As New List(Of String)()
            For Each value In values
                Dim text = If(value, String.Empty)
                text = text.Replace(vbCrLf, " ")
                text = text.Replace(vbCr, " ")
                text = text.Replace(vbLf, " ")
                text = text.Replace(vbTab, " ")

                Dim firstSignificantIndex As Integer = 0
                While firstSignificantIndex < text.Length AndAlso text(firstSignificantIndex) = " "c
                    firstSignificantIndex += 1
                End While

                If firstSignificantIndex < text.Length Then
                    Select Case text(firstSignificantIndex)
                        Case "="c, "+"c, "-"c, "@"c
                            text = "'" & text
                    End Select
                End If

                text = text.Replace("""", """""")
                escaped.Add("""" & text & """")
            Next
            Return escaped
        End Function

        Private Sub CopyToClipboard(text As String)
            If String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            Try
                Clipboard.SetText(text)
                SetStatus("Copied to clipboard.", False)
            Catch ex As Exception
                SetStatus("Clipboard is unavailable.", True)
            End Try
        End Sub

        Private Sub SetStatus(message As String, isError As Boolean)
            _statusLabel.Text = message
            _statusLabel.ForeColor = If(isError, Color.Firebrick, Color.DarkGreen)
        End Sub
    End Class
End Namespace
