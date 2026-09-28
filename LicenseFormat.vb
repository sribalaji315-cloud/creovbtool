Imports System
Imports System.Collections.Generic
Imports System.Security.Cryptography
Imports System.Text

Namespace CreoRegenAppVB

    Friend NotInheritable Class LicenseInfo
        Friend Property LicenseId As String
        Friend Property Licensee As String
        Friend Property IssuedUtc As Date
    End Class

    ''' <summary>
    ''' Wire format for BOMForge license keys. Compiled into both the application
    ''' (verify only) and the key generator (sign + verify).
    ''' </summary>
    Friend NotInheritable Class LicenseFormat

        Friend Const InvalidKeyMessage As String = "This license key is not valid."
        Friend Const FormatVersion As Byte = 1
        Friend Const SignatureLength As Integer = 64
        Friend Const LicenseIdLength As Integer = 8
        Friend Const MaxLicenseeLength As Integer = 64
        Friend Const KeyGroupSize As Integer = 5

        ' Public half of the ECDSA P-256 signing key, as a Base64 CNG EccPublicBlob.
        ' Replace with the value printed by BOMForgeKeyGen on its first run.
        Friend Const PublicKeyBase64 As String = "RUNTMSAAAADketJQ0puU5XwSnpodEhZnNGgC+SkubAvzQbEtnJDHUxewDYE6JbmBF2sCOyqOSmkAB9nqYosyuF2AdJWXiCKY"

        ' version(1) + licenseId(8) + issueDays(2) + licenseeLength(1)
        Private Const HeaderLength As Integer = 12

        ' Crockford Base32: no I, L, O or U.
        Private Const Alphabet As String = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"

        Private Shared ReadOnly Epoch As Date = New Date(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc)

        Private Sub New()
        End Sub

        Friend Shared Function NormalizeLicensee(licensee As String) As String
            If String.IsNullOrWhiteSpace(licensee) Then
                Return Nothing
            End If

            Dim trimmed = licensee.Trim()
            If trimmed.Length > MaxLicenseeLength Then
                Return Nothing
            End If

            For Each ch In trimmed
                If Char.IsControl(ch) Then
                    Return Nothing
                End If
            Next

            Return trimmed
        End Function

        Friend Shared Function BuildPayload(licenseId As Byte(), licensee As String, issuedUtc As Date) As Byte()
            If licenseId Is Nothing OrElse licenseId.Length <> LicenseIdLength Then
                Throw New ArgumentException("License id must be " & LicenseIdLength.ToString() & " bytes.")
            End If

            Dim normalizedName = NormalizeLicensee(licensee)
            If normalizedName Is Nothing Then
                Throw New ArgumentException("Licensee name is invalid.")
            End If

            Dim nameBytes = Encoding.UTF8.GetBytes(normalizedName)
            If nameBytes.Length < 1 OrElse nameBytes.Length > 255 Then
                Throw New ArgumentException("Licensee name is too long.")
            End If

            Dim issuedDays = CInt(Math.Floor((issuedUtc.Date - Epoch.Date).TotalDays))
            If issuedDays < 0 OrElse issuedDays > UShort.MaxValue Then
                Throw New ArgumentException("Issue date is out of range.")
            End If

            Dim payload(HeaderLength + nameBytes.Length - 1) As Byte
            payload(0) = FormatVersion
            Array.Copy(licenseId, 0, payload, 1, LicenseIdLength)
            payload(9) = CByte((issuedDays >> 8) And &HFF)
            payload(10) = CByte(issuedDays And &HFF)
            payload(11) = CByte(nameBytes.Length)
            Array.Copy(nameBytes, 0, payload, HeaderLength, nameBytes.Length)
            Return payload
        End Function

        Friend Shared Function ComposeKey(payload As Byte(), signature As Byte()) As String
            If payload Is Nothing OrElse payload.Length < HeaderLength Then
                Throw New ArgumentException("Payload is invalid.")
            End If

            If signature Is Nothing OrElse signature.Length <> SignatureLength Then
                Throw New ArgumentException("Signature is invalid.")
            End If

            Dim blob(payload.Length + SignatureLength - 1) As Byte
            Array.Copy(payload, 0, blob, 0, payload.Length)
            Array.Copy(signature, 0, blob, payload.Length, SignatureLength)
            Return GroupKeyText(EncodeBase32(blob))
        End Function

        Friend Shared Function TryParse(keyText As String, ByRef info As LicenseInfo) As Boolean
            info = Nothing

            Dim blob As Byte() = Nothing
            If Not TryDecodeBase32(keyText, blob) Then
                Return False
            End If

            If blob.Length < HeaderLength + 1 + SignatureLength Then
                Return False
            End If

            Dim payloadLength = blob.Length - SignatureLength
            Dim payload(payloadLength - 1) As Byte
            Dim signature(SignatureLength - 1) As Byte
            Array.Copy(blob, 0, payload, 0, payloadLength)
            Array.Copy(blob, payloadLength, signature, 0, SignatureLength)

            Dim parsed As LicenseInfo = Nothing
            If Not TryReadPayload(payload, parsed) Then
                Return False
            End If

            If Not VerifySignature(payload, signature) Then
                Return False
            End If

            info = parsed
            Return True
        End Function

        Private Shared Function TryReadPayload(payload As Byte(), ByRef info As LicenseInfo) As Boolean
            info = Nothing

            If payload(0) <> FormatVersion Then
                Return False
            End If

            Dim nameLength = CInt(payload(11))
            If nameLength < 1 OrElse HeaderLength + nameLength <> payload.Length Then
                Return False
            End If

            Dim licensee As String
            Try
                licensee = New UTF8Encoding(False, True).GetString(payload, HeaderLength, nameLength)
            Catch
                Return False
            End Try

            Dim normalizedName = NormalizeLicensee(licensee)
            If normalizedName Is Nothing Then
                Return False
            End If

            Dim issuedDays = (CInt(payload(9)) << 8) Or CInt(payload(10))

            Dim idBuilder As New StringBuilder(LicenseIdLength * 2)
            For index = 1 To LicenseIdLength
                idBuilder.Append(payload(index).ToString("X2"))
            Next

            info = New LicenseInfo With {
                .LicenseId = idBuilder.ToString(),
                .Licensee = normalizedName,
                .IssuedUtc = Epoch.AddDays(issuedDays)
            }
            Return True
        End Function

        Private Shared Function VerifySignature(payload As Byte(), signature As Byte()) As Boolean
            Dim publicBlob As Byte()
            Try
                publicBlob = Convert.FromBase64String(PublicKeyBase64)
            Catch
                Return False
            End Try

            Try
                Using publicKey = CngKey.Import(publicBlob, CngKeyBlobFormat.EccPublicBlob)
                    Using verifier As New ECDsaCng(publicKey)
                        verifier.HashAlgorithm = CngAlgorithm.Sha256
                        Return verifier.VerifyData(payload, signature)
                    End Using
                End Using
            Catch
                Return False
            End Try
        End Function

        Friend Shared Function GroupKeyText(rawKey As String) As String
            If String.IsNullOrEmpty(rawKey) Then
                Return String.Empty
            End If

            Dim builder As New StringBuilder(rawKey.Length + rawKey.Length \ KeyGroupSize)
            For index = 0 To rawKey.Length - 1
                If index > 0 AndAlso index Mod KeyGroupSize = 0 Then
                    builder.Append("-"c)
                End If
                builder.Append(rawKey(index))
            Next
            Return builder.ToString()
        End Function

        Private Shared Function EncodeBase32(data As Byte()) As String
            Dim builder As New StringBuilder(((data.Length * 8) + 4) \ 5)
            Dim buffer As Integer = 0
            Dim bitsHeld As Integer = 0

            For Each value In data
                buffer = (buffer << 8) Or CInt(value)
                bitsHeld += 8
                While bitsHeld >= 5
                    builder.Append(Alphabet.Chars((buffer >> (bitsHeld - 5)) And 31))
                    bitsHeld -= 5
                End While
            Next

            If bitsHeld > 0 Then
                builder.Append(Alphabet.Chars((buffer << (5 - bitsHeld)) And 31))
            End If

            Return builder.ToString()
        End Function

        Private Shared Function TryDecodeBase32(text As String, ByRef data As Byte()) As Boolean
            data = Nothing

            If String.IsNullOrWhiteSpace(text) Then
                Return False
            End If

            ' Guards against a pathological paste being decoded at all.
            If text.Length > 4096 Then
                Return False
            End If

            Dim bytes As New List(Of Byte)()
            Dim buffer As Integer = 0
            Dim bitsHeld As Integer = 0

            For Each ch In text
                If ch = "-"c OrElse Char.IsWhiteSpace(ch) Then
                    Continue For
                End If

                Dim symbol = Char.ToUpperInvariant(ch)
                Select Case symbol
                    Case "O"c
                        symbol = "0"c
                    Case "I"c, "L"c
                        symbol = "1"c
                End Select

                Dim value = Alphabet.IndexOf(symbol)
                If value < 0 Then
                    Return False
                End If

                buffer = (buffer << 5) Or value
                bitsHeld += 5
                If bitsHeld >= 8 Then
                    bytes.Add(CByte((buffer >> (bitsHeld - 8)) And &HFF))
                    bitsHeld -= 8
                End If
            Next

            If bitsHeld >= 5 Then
                Return False
            End If

            If bitsHeld > 0 AndAlso (buffer And ((1 << bitsHeld) - 1)) <> 0 Then
                Return False
            End If

            If bytes.Count = 0 Then
                Return False
            End If

            data = bytes.ToArray()
            Return True
        End Function
    End Class
End Namespace
