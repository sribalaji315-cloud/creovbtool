Imports System

Namespace CreoRegenAppVB
    Friend Enum CreoModelKind
        Unknown = 0
        Part = 1
        Assembly = 2
        Drawing = 4
    End Enum

    Friend Class ParsedModelName
        Friend Property Name As String
        Friend Property Kind As CreoModelKind
        Friend Property Ext As String
        Friend Property FileName As String
        Friend Property HasExplicitExtension As Boolean
    End Class

    Friend NotInheritable Class ModelNameParser
        Private Const MaxNameLength As Integer = 260

        Private Shared ReadOnly AllowedExtensions As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "ASM", "DRW", "PRT", "MFG", "FRM", "LAY", "DGM", "SEC", "SKF"
        }

        Private Sub New()
        End Sub

        Friend Shared Function Parse(rawInput As String) As ParsedModelName
            If String.IsNullOrWhiteSpace(rawInput) Then
                Throw New ArgumentException("Model name is empty.")
            End If

            Dim trimmed = rawInput.Trim()
            If trimmed.Length > MaxNameLength Then
                Throw New ArgumentException("Model name exceeds maximum length.")
            End If

            Dim versionless = StripVersionSuffix(trimmed)
            Dim dot = versionless.LastIndexOf("."c)
            If dot <= 0 Then
                Return New ParsedModelName With {
                    .Name = versionless,
                    .Kind = CreoModelKind.Unknown,
                    .Ext = String.Empty,
                    .FileName = versionless,
                    .HasExplicitExtension = False
                }
            End If

            Dim modelName = versionless.Substring(0, dot)
            Dim ext = versionless.Substring(dot + 1).Trim().ToUpperInvariant()
            If modelName.Length = 0 Then
                Throw New ArgumentException("Model name is empty.")
            End If

            If Not AllowedExtensions.Contains(ext) Then
                Throw New ArgumentException("Unsupported model extension: ." & ext.ToLowerInvariant())
            End If

            Select Case ext
                Case "ASM"
                    Return New ParsedModelName With {
                        .Name = modelName,
                        .Kind = CreoModelKind.Assembly,
                        .Ext = ext,
                        .FileName = modelName & "." & ext.ToLowerInvariant(),
                        .HasExplicitExtension = True
                    }
                Case "DRW"
                    Return New ParsedModelName With {
                        .Name = modelName,
                        .Kind = CreoModelKind.Drawing,
                        .Ext = ext,
                        .FileName = modelName & "." & ext.ToLowerInvariant(),
                        .HasExplicitExtension = True
                    }
                Case "PRT"
                    Return New ParsedModelName With {
                        .Name = modelName,
                        .Kind = CreoModelKind.Part,
                        .Ext = ext,
                        .FileName = modelName & "." & ext.ToLowerInvariant(),
                        .HasExplicitExtension = True
                    }
                Case Else
                    Return New ParsedModelName With {
                        .Name = modelName,
                        .Kind = CreoModelKind.Unknown,
                        .Ext = ext,
                        .FileName = modelName & "." & ext.ToLowerInvariant(),
                        .HasExplicitExtension = True
                    }
            End Select
        End Function

        Private Shared Function StripVersionSuffix(inputText As String) As String
            Dim lastDot = inputText.LastIndexOf("."c)
            If lastDot <= 0 OrElse lastDot = inputText.Length - 1 Then
                Return inputText
            End If

            Dim suffix = inputText.Substring(lastDot + 1)
            If suffix.Length = 0 Then
                Return inputText
            End If

            For Each ch In suffix
                If Not Char.IsDigit(ch) Then
                    Return inputText
                End If
            Next

            Return inputText.Substring(0, lastDot)
        End Function
    End Class
End Namespace
