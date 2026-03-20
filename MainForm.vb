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
        Private ReadOnly _loadParametersButton As Button
        Private ReadOnly _setParameterButton As Button
        Private ReadOnly _clearLogButton As Button
        Private ReadOnly _outputTextBox As TextBox
        Private ReadOnly _tabs As TabControl
        Private ReadOnly _paramsGrid As DataGridView

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
            top.ColumnCount = 8
            top.RowCount = 2
            top.AutoSize = True
            top.AutoSizeMode = AutoSizeMode.GrowAndShrink
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            top.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
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

            _loadParametersButton = New Button()
            _loadParametersButton.Text = "Load Parameters"
            AddHandler _loadParametersButton.Click, AddressOf LoadParametersClickedAsync

            _setParameterButton = New Button()
            _setParameterButton.Text = "Set Selected"
            AddHandler _setParameterButton.Click, AddressOf SetSelectedParameterClickedAsync

            _clearLogButton = New Button()
            _clearLogButton.Text = "Clear Log"
            AddHandler _clearLogButton.Click, Sub(sender, args) _outputTextBox.Clear()

            top.Controls.Add(_modelLabel, 0, 0)
            top.Controls.Add(_modelNameTextBox, 1, 0)
            top.SetColumnSpan(_modelNameTextBox, 5)

            top.Controls.Add(_connectButton, 2, 1)
            top.Controls.Add(_retrieveButton, 3, 1)
            top.Controls.Add(_regenerateButton, 4, 1)
            top.Controls.Add(_loadParametersButton, 5, 1)
            top.Controls.Add(_setParameterButton, 6, 1)
            top.Controls.Add(_clearLogButton, 7, 1)

            _outputTextBox = New TextBox()
            _outputTextBox.Multiline = True
            _outputTextBox.ScrollBars = ScrollBars.Vertical
            _outputTextBox.ReadOnly = True
            _outputTextBox.Dock = DockStyle.Fill
            _outputTextBox.BackColor = Color.FromArgb(30, 30, 30)
            _outputTextBox.ForeColor = Color.FromArgb(200, 255, 200)
            _outputTextBox.Font = New Font(FontFamily.GenericMonospace, 9.0F)

            _paramsGrid = New DataGridView()
            _paramsGrid.Dock = DockStyle.Fill
            _paramsGrid.AllowUserToAddRows = False
            _paramsGrid.AllowUserToDeleteRows = False
            _paramsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            _paramsGrid.MultiSelect = False
            _paramsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            _paramsGrid.Columns.Add("Name", "Parameter")
            _paramsGrid.Columns.Add("Type", "Type")
            _paramsGrid.Columns.Add("Value", "Current Value")
            _paramsGrid.Columns.Add("NewValue", "New Value")
            _paramsGrid.Columns.Add("Description", "Description")
            Dim discrColumn = New DataGridViewTextBoxColumn()
            discrColumn.Name = "Discr"
            discrColumn.Visible = False
            _paramsGrid.Columns.Add(discrColumn)
            _paramsGrid.Columns("Name").ReadOnly = True
            _paramsGrid.Columns("Type").ReadOnly = True
            _paramsGrid.Columns("Value").ReadOnly = True
            _paramsGrid.Columns("Description").ReadOnly = True

            _tabs = New TabControl()
            _tabs.Dock = DockStyle.Fill

            Dim logTab As New TabPage("Log")
            logTab.Controls.Add(_outputTextBox)

            Dim paramsTab As New TabPage("Parameters")
            paramsTab.Controls.Add(_paramsGrid)

            _tabs.TabPages.Add(logTab)
            _tabs.TabPages.Add(paramsTab)

            root.Controls.Add(top, 0, 0)
            root.Controls.Add(_tabs, 0, 1)

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

        Private Async Sub LoadParametersClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            SetBusyState(True)
            Try
                AddLog("Load Parameters started...")
                Dim result = Await _service.GetParametersAsync(modelInput)
                If Not result.Success Then
                    AddLog("ERROR: " & result.Message)
                    Return
                End If

                _paramsGrid.Rows.Clear()
                For Each item In result.Items
                    Dim typeName = GetDiscrName(item.Discr)
                    _paramsGrid.Rows.Add(item.Name, typeName, item.ValueText, "", item.Description, item.Discr)
                Next

                _tabs.SelectedIndex = 1
                AddLog("OK: " & result.Message)
            Catch ex As Exception
                AddLog("ERROR: " & ex.Message)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Async Sub SetSelectedParameterClickedAsync(sender As Object, e As EventArgs)
            If _paramsGrid.SelectedRows.Count = 0 Then
                AddLog("ERROR: Select a parameter row first.")
                Return
            End If

            Dim row = _paramsGrid.SelectedRows(0)
            Dim parameterName = Convert.ToString(row.Cells("Name").Value)
            Dim userInput = Convert.ToString(row.Cells("NewValue").Value)
            
            If String.IsNullOrWhiteSpace(userInput) Then
                AddLog("ERROR: Please type a value in the 'New Value' column for " & parameterName & " before setting.")
                Return
            End If

            Dim modelInput = _modelNameTextBox.Text

            Await RunOperationAsync(Function() _service.SetParameterAsync(modelInput, parameterName, userInput), "Set Parameter")
            
            ' Clear the new value column if successful
            row.Cells("NewValue").Value = ""
            
            ' We might also want to reload parameters to show the updated current value, but saving manual click for now.
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
            _loadParametersButton.Enabled = Not isBusy
            _setParameterButton.Enabled = Not isBusy
            _clearLogButton.Enabled = Not isBusy
        End Sub

        Private Function GetDiscrName(discr As Integer) As String
            Select Case discr
                Case 0
                    Return "String"
                Case 1
                    Return "Integer"
                Case 2
                    Return "Boolean"
                Case 3
                    Return "Double"
                Case 4
                    Return "NoteId"
                Case Else
                    Return "Type " & discr.ToString()
            End Select
        End Function

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
