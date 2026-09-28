Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace CreoRegenAppVB

    Friend Class SplashForm
        Inherits Form

        Private Const DisplayMilliseconds As Integer = 1500

        Private ReadOnly _timer As Timer

        Friend Sub New(licensee As String)
            FormBorderStyle = FormBorderStyle.None
            StartPosition = FormStartPosition.CenterScreen
            ShowInTaskbar = False
            TopMost = True
            Size = New Size(380, 130)
            BackColor = Color.FromArgb(30, 30, 30)

            Dim layout As New TableLayoutPanel()
            layout.Dock = DockStyle.Fill
            layout.ColumnCount = 1
            layout.RowCount = 2
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 55.0F))
            layout.RowStyles.Add(New RowStyle(SizeType.Percent, 45.0F))

            Dim titleLabel As New Label()
            titleLabel.Text = "BOMForge"
            titleLabel.Dock = DockStyle.Fill
            titleLabel.TextAlign = ContentAlignment.BottomCenter
            titleLabel.ForeColor = Color.White
            titleLabel.Font = New Font(Font.FontFamily, 20.0F, FontStyle.Bold)

            Dim licenseeLabel As New Label()
            licenseeLabel.Text = "Licensed to " & If(String.IsNullOrWhiteSpace(licensee), "—", licensee)
            licenseeLabel.Dock = DockStyle.Fill
            licenseeLabel.TextAlign = ContentAlignment.TopCenter
            licenseeLabel.ForeColor = Color.FromArgb(200, 255, 200)
            licenseeLabel.Padding = New Padding(8, 6, 8, 0)

            layout.Controls.Add(titleLabel, 0, 0)
            layout.Controls.Add(licenseeLabel, 0, 1)
            Controls.Add(layout)

            AddHandler Click, AddressOf DismissRequested
            AddHandler titleLabel.Click, AddressOf DismissRequested
            AddHandler licenseeLabel.Click, AddressOf DismissRequested

            _timer = New Timer()
            _timer.Interval = DisplayMilliseconds
            AddHandler _timer.Tick, AddressOf DismissRequested
        End Sub

        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            _timer.Start()
        End Sub

        Private Sub DismissRequested(sender As Object, e As EventArgs)
            _timer.Stop()
            Close()
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing AndAlso _timer IsNot Nothing Then
                _timer.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub
    End Class
End Namespace
