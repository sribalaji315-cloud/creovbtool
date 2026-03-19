Imports System

Namespace CreoRegenAppVB
    Friend Enum CreoModelKind
        Part = 1
        Assembly = 2
        Drawing = 4
    End Enum

    Friend Class ParsedModelName
        Friend Property Name As String
        Friend Property Kind As CreoModelKind
        Friend Property Ext As String
        Friend Property FileName As String
    End Class

    Friend NotInheritable Class ModelNameParser
        Private Sub New()
        End Sub

        Friend Shared Function Parse(rawInput As String) As ParsedModelName
            If String.IsNullOrWhiteSpace(rawInput) Then
                Throw New ArgumentException("Model name is empty.")
            End If

            Dim trimmed = rawInput.Trim()
            Dim dot = trimmed.LastIndexOf("."c)
            If dot <= 0 Then
                Return New ParsedModelName With {
                    .Name = trimmed,
                    .Kind = CreoModelKind.Part,
                    .Ext = "PRT",
                    .FileName = trimmed & ".prt"
                }
            End If

            Dim modelName = trimmed.Substring(0, dot)
            Dim ext = trimmed.Substring(dot + 1).Trim().ToUpperInvariant()
            If modelName.Length = 0 Then
                Throw New ArgumentException("Model name is empty.")
            End If

            Select Case ext
                Case "ASM"
                    Return New ParsedModelName With {
                        .Name = modelName,
                        .Kind = CreoModelKind.Assembly,
                        .Ext = ext,
                        .FileName = modelName & "." & ext.ToLowerInvariant()
                    }
                Case "DRW"
                    Return New ParsedModelName With {
                        .Name = modelName,
                        .Kind = CreoModelKind.Drawing,
                        .Ext = ext,
                        .FileName = modelName & "." & ext.ToLowerInvariant()
                    }
                Case Else
                    Return New ParsedModelName With {
                        .Name = modelName,
                        .Kind = CreoModelKind.Part,
                        .Ext = ext,
                        .FileName = modelName & "." & ext.ToLowerInvariant()
                    }
            End Select
        End Function
    End Class
End Namespace
