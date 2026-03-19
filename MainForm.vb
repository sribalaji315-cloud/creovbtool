Imports System
Imports System.Drawing
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace CreoRegenAppVB
    Friend Class MainForm
        Inherits Form

        Private ReadOnly _service As CreoSessionService

        Private ReadOnly _modelLabel As Label
        Private ReadOnly _modelNameTextBox As TextBox
        Private ReadOnly _connectButton As Button
        Private ReadOnly _retrieveButton As Button
        Private ReadOnly _regenerateButton As Button
        Private ReadOnly _clearLogButton As Button
        Private ReadOnly _outputTextBox As TextBox

        Friend Sub New()
            Text = "Creo Regen Toolkit (VB.NET)"
            Width = 760
            Height = 480
            MinimumSize = New Size(560, 360)
            StartPosition = FormStartPosition.CenterScreen

            _service = New CreoSessionService()

            Dim root As New TableLayoutPanel()
            root.Dock = DockStyle.Fill
            root.Padding = New Padding(10)
            root.ColumnCount = 1
            root.RowCount = 2
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim top As New TableLayoutPanel()
            top.Dock = DockStyle.Top
            top.ColumnCount = 6
            top.RowCount = 2
            top.AutoSize = True
            top.AutoSizeMode = AutoSizeMode.GrowAndShrink
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            _modelLabel = New Label()
            _modelLabel.Text = "Model name:"
            _modelLabel.AutoSize = True
            _modelLabel.Anchor = AnchorStyles.Left

            _modelNameTextBox = New TextBox()
            _modelNameTextBox.Width = 260
            _modelNameTextBox.Anchor = AnchorStyles.Left Or AnchorStyles.Right

            _connectButton = New Button()
            _connectButton.Text = "Connect"
            AddHandler _connectButton.Click, AddressOf ConnectClickedAsync

            _retrieveButton = New Button()
            _retrieveButton.Text = "Retrieve"
            AddHandler _retrieveButton.Click, AddressOf RetrieveClickedAsync

            _regenerateButton = New Button()
            _regenerateButton.Text = "Regenerate"
            AddHandler _regenerateButton.Click, AddressOf RegenerateClickedAsync

            _clearLogButton = New Button()
            _clearLogButton.Text = "Clear Log"
            AddHandler _clearLogButton.Click, Sub(sender, args) _outputTextBox.Clear()

            top.Controls.Add(_modelLabel, 0, 0)
            top.Controls.Add(_modelNameTextBox, 1, 0)
            top.SetColumnSpan(_modelNameTextBox, 5)

            top.Controls.Add(_connectButton, 2, 1)
            top.Controls.Add(_retrieveButton, 3, 1)
            top.Controls.Add(_regenerateButton, 4, 1)
            top.Controls.Add(_clearLogButton, 5, 1)

            _outputTextBox = New TextBox()
            _outputTextBox.Multiline = True
            _outputTextBox.ScrollBars = ScrollBars.Vertical
            _outputTextBox.ReadOnly = True
            _outputTextBox.Dock = DockStyle.Fill
            _outputTextBox.BackColor = Color.FromArgb(30, 30, 30)
            _outputTextBox.ForeColor = Color.FromArgb(200, 255, 200)
            _outputTextBox.Font = New Font(FontFamily.GenericMonospace, 9.0F)

            root.Controls.Add(top, 0, 0)
            root.Controls.Add(_outputTextBox, 0, 1)

            Controls.Add(root)

            AddLog("UI ready.")
            AddLog("Example: B00001341 or B00001341.PRT or myassy.ASM")
        End Sub

        Private Async Sub ConnectClickedAsync(sender As Object, e As EventArgs)
            Await RunOperationAsync(
                Async Function()
                    Dim settings = New CreoConnectionSettings With {
                        .CreoStartCommand = "parametric",
                        .TextDirectory = "."
                    }
                    Return Await _service.EnsureConnectedAsync(settings)
                End Function,
                "Connect")
        End Sub

        Private Async Sub RetrieveClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            Await RunOperationAsync(Function() _service.RetrieveModelAsync(modelInput), "Retrieve")
        End Sub

        Private Async Sub RegenerateClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            Await RunOperationAsync(Function() _service.RegenerateModelAsync(modelInput), "Regenerate")
        End Sub

        Private Async Function RunOperationAsync(workFactory As Func(Of Task(Of OperationResult)), actionName As String) As Task
            SetBusyState(True)
            Try
                AddLog(actionName & " started...")
                Dim result = Await workFactory()
                If result.Success Then
                    AddLog("OK: " & result.Message)
                Else
                    AddLog("ERROR: " & result.Message)
                End If
            Catch ex As Exception
                AddLog("ERROR: " & ex.Message)
            Finally
                SetBusyState(False)
            End Try
        End Function

        Private Sub SetBusyState(isBusy As Boolean)
            _connectButton.Enabled = Not isBusy
            _retrieveButton.Enabled = Not isBusy
            _regenerateButton.Enabled = Not isBusy
            _clearLogButton.Enabled = Not isBusy
        End Sub

        Private Sub AddLog(message As String)
            Dim line = String.Format("[{0:HH:mm:ss}] {1}", Date.Now, message)
            _outputTextBox.AppendText(line & Environment.NewLine)
            _outputTextBox.SelectionStart = _outputTextBox.TextLength
            _outputTextBox.ScrollToCaret()
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            MyBase.OnFormClosing(e)

            Try
                _service.Dispose()
            Catch
            End Try
        End Sub
    End Class
End Namespace
