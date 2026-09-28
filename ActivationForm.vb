Imports System
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Namespace CreoRegenAppVB

    Friend Class ActivationForm
        Inherits Form

        Private Const MaxKeyFileBytes As Integer = 8192

        Private ReadOnly _status As LicenseStatus
        Private ReadOnly _blocking As Boolean
        Private ReadOnly _statusLabel As Label
        Private ReadOnly _messageLabel As Label
        Private ReadOnly _keyTextBox As TextBox
        Private ReadOnly _activateButton As Button
        Private ReadOnly _loadFileButton As Button
        Private ReadOnly _dismissButton As Button

        Private _resultStatus As LicenseStatus

        Friend Sub New(status As LicenseStatus)
            _status = status
            _blocking = status.State = LicenseState.Expired
            _resultStatus = status

            Text = "BOMForge Activation"
            Width = 520
            Height = 360
            FormBorderStyle = FormBorderStyle.FixedDialog
            MaximizeBox = False
            MinimizeBox = False
            StartPosition = FormStartPosition.CenterScreen
            ShowInTaskbar = True

            Dim root As New TableLayoutPanel()
            root.Dock = DockStyle.Fill
            root.Padding = New Padding(14)
            root.ColumnCount = 1
            root.RowCount = 6
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            Dim titleLabel As New Label()
            titleLabel.Text = "BOMForge"
            titleLabel.AutoSize = True
            titleLabel.Font = New Font(Font.FontFamily, 14.0F, FontStyle.Bold)
            titleLabel.Margin = New Padding(0, 0, 0, 6)

            _statusLabel = New Label()
            _statusLabel.AutoSize = True
            _statusLabel.MaximumSize = New Size(460, 0)
            _statusLabel.Text = BuildStatusText()
            _statusLabel.Margin = New Padding(0, 0, 0, 10)

            Dim promptLabel As New Label()
            promptLabel.AutoSize = True
            promptLabel.MaximumSize = New Size(460, 0)
            promptLabel.Text = "Paste your license key below, or load a .lic file."
            promptLabel.Margin = New Padding(0, 0, 0, 4)

            _keyTextBox = New TextBox()
            _keyTextBox.Multiline = True
            _keyTextBox.ScrollBars = ScrollBars.Vertical
            _keyTextBox.Dock = DockStyle.Fill
            _keyTextBox.MaxLength = MaxKeyFileBytes
            _keyTextBox.Font = New Font(FontFamily.GenericMonospace, 9.0F)

            _messageLabel = New Label()
            _messageLabel.AutoSize = True
            _messageLabel.MaximumSize = New Size(460, 0)
            _messageLabel.ForeColor = Color.Firebrick
            _messageLabel.Margin = New Padding(0, 8, 0, 4)

            Dim buttonRow As New FlowLayoutPanel()
            buttonRow.Dock = DockStyle.Fill
            buttonRow.FlowDirection = FlowDirection.RightToLeft
            buttonRow.AutoSize = True
            buttonRow.AutoSizeMode = AutoSizeMode.GrowAndShrink

            _dismissButton = New Button()
            _dismissButton.Text = If(_blocking, "Exit", "Continue Trial")
            _dismissButton.AutoSize = True
            _dismissButton.Margin = New Padding(4)
            AddHandler _dismissButton.Click, AddressOf DismissClicked

            _activateButton = New Button()
            _activateButton.Text = "Activate"
            _activateButton.AutoSize = True
            _activateButton.Margin = New Padding(4)
            AddHandler _activateButton.Click, AddressOf ActivateClicked

            _loadFileButton = New Button()
            _loadFileButton.Text = "Load .lic File..."
            _loadFileButton.AutoSize = True
            _loadFileButton.Margin = New Padding(4)
            AddHandler _loadFileButton.Click, AddressOf LoadFileClicked

            buttonRow.Controls.Add(_dismissButton)
            buttonRow.Controls.Add(_activateButton)
            buttonRow.Controls.Add(_loadFileButton)

            root.Controls.Add(titleLabel, 0, 0)
            root.Controls.Add(_statusLabel, 0, 1)
            root.Controls.Add(promptLabel, 0, 2)
            root.Controls.Add(_keyTextBox, 0, 3)
            root.Controls.Add(_messageLabel, 0, 4)
            root.Controls.Add(buttonRow, 0, 5)

            Controls.Add(root)

            AcceptButton = _activateButton
            If Not _blocking Then
                CancelButton = _dismissButton
            End If
        End Sub

        Friend ReadOnly Property ResultStatus As LicenseStatus
            Get
                Return _resultStatus
            End Get
        End Property

        Private Function BuildStatusText() As String
            Select Case _status.State
                Case LicenseState.Licensed
                    Return "This copy is licensed to " & _status.Licensee & "."
                Case LicenseState.Trial
                    Return "Trial mode — " & _status.DaysRemaining.ToString() &
                           If(_status.DaysRemaining = 1, " day", " days") & " remaining of " &
                           LicenseManager.TrialLengthDays.ToString() & "."
                Case Else
                    Return "Your " & LicenseManager.TrialLengthDays.ToString() &
                           "-day trial has ended. Enter a license key to continue using BOMForge."
            End Select
        End Function

        Private Sub ActivateClicked(sender As Object, e As EventArgs)
            _messageLabel.Text = String.Empty

            Dim keyText = _keyTextBox.Text
            If String.IsNullOrWhiteSpace(keyText) Then
                _messageLabel.Text = "Enter a license key."
                Return
            End If

            Dim info As LicenseInfo = Nothing
            If Not LicenseManager.TryActivate(keyText, info) Then
                _messageLabel.Text = LicenseFormat.InvalidKeyMessage
                Return
            End If

            _resultStatus = New LicenseStatus With {
                .State = LicenseState.Licensed,
                .DaysRemaining = 0,
                .Licensee = info.Licensee
            }

            MessageBox.Show(Me, "BOMForge is now licensed to " & info.Licensee & ".",
                            "Activation Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

            DialogResult = DialogResult.OK
            Close()
        End Sub

        Private Sub LoadFileClicked(sender As Object, e As EventArgs)
            _messageLabel.Text = String.Empty

            Using dialog As New OpenFileDialog()
                dialog.Filter = "License files (*.lic)|*.lic|Text files (*.txt)|*.txt|All files (*.*)|*.*"
                dialog.Title = "Select BOMForge License File"
                dialog.CheckFileExists = True

                If dialog.ShowDialog(Me) <> DialogResult.OK Then
                    Return
                End If

                Try
                    Dim info As New FileInfo(dialog.FileName)
                    If info.Length > MaxKeyFileBytes Then
                        _messageLabel.Text = "That file is not a license file."
                        Return
                    End If

                    _keyTextBox.Text = File.ReadAllText(dialog.FileName).Trim()
                Catch ex As Exception
                    CreoSessionService.LogApplicationError("Load license file", ex)
                    _messageLabel.Text = "The license file could not be read."
                End Try
            End Using
        End Sub

        Private Sub DismissClicked(sender As Object, e As EventArgs)
            DialogResult = If(_blocking, DialogResult.Cancel, DialogResult.OK)
            Close()
        End Sub
    End Class
End Namespace
