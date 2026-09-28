Imports System
Imports System.Windows.Forms

Namespace CreoRegenAppVB
    Friend Module Program
        <STAThread>
        Public Sub Main()
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)

            Dim status = LicenseManager.Evaluate()

            If status.State <> LicenseState.Licensed Then
                Using dialog As New ActivationForm(status)
                    If dialog.ShowDialog() <> DialogResult.OK Then
                        Return
                    End If

                    status = dialog.ResultStatus
                End Using
            End If

            If status.State = LicenseState.Expired Then
                Return
            End If

            If status.State = LicenseState.Licensed Then
                Using splash As New SplashForm(status.Licensee)
                    splash.ShowDialog()
                End Using
            End If

            Application.Run(New MainForm(status))
        End Sub
    End Module
End Namespace
