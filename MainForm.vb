Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace CreoRegenAppVB
    Friend Class MainForm
        Inherits Form

        Private Const BomTabIndex As Integer = 2
        Private Const ItemPanelMinWidth As Integer = 320

        Private ReadOnly _service As CreoSessionService

        Private ReadOnly _modelLabel As Label
        Private ReadOnly _modelNameTextBox As TextBox
        Private ReadOnly _connectButton As Button
        Private ReadOnly _retrieveButton As Button
        Private ReadOnly _regenerateButton As Button
        Private ReadOnly _loadParametersButton As Button
        Private ReadOnly _loadFeaturesButton As Button
        Private ReadOnly _loadComponentsButton As Button
        Private ReadOnly _buildBomButton As Button
        Private ReadOnly _exportBomButton As Button
        Private ReadOnly _expandAllBomButton As Button
        Private ReadOnly _collapseAllBomButton As Button
        Private ReadOnly _showItemParametersButton As Button
        Private ReadOnly _setParameterButton As Button
        Private ReadOnly _clearLogButton As Button
        Private ReadOnly _outputTextBox As TextBox
        Private ReadOnly _tabs As TabControl
        Private ReadOnly _paramsGrid As DataGridView
        Private ReadOnly _featuresGrid As DataGridView
        Private ReadOnly _componentsGrid As DataGridView
        Private ReadOnly _bomGrid As DataGridView
        Private ReadOnly _bomRowsAll As List(Of BomGridRowItem)
        Private ReadOnly _bomRowsByNodeId As Dictionary(Of Integer, BomGridRowItem)
        Private ReadOnly _bomCollapsedNodeIds As HashSet(Of Integer)
        Private ReadOnly _itemParamsPanel As Panel
        Private ReadOnly _itemParamsTitle As Label
        Private ReadOnly _itemParamsLoadingBar As ProgressBar
        Private ReadOnly _itemParamsGrid As DataGridView
        Private ReadOnly _itemSetParameterButton As Button
        Private _selectedBomModelInput As String
        Private _suppressSelectionChanged As Boolean
        Private _isLoadingItemParams As Boolean

        Friend Sub New()
            Text = "Creo Regen Toolkit (VB.NET)"
            Width = 760
            Height = 480
            MinimumSize = New Size(560, 360)
            StartPosition = FormStartPosition.CenterScreen
            SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
            UpdateStyles()

            _service = New CreoSessionService()
            _bomRowsAll = New List(Of BomGridRowItem)()
            _bomRowsByNodeId = New Dictionary(Of Integer, BomGridRowItem)()
            _bomCollapsedNodeIds = New HashSet(Of Integer)()
            Dim root As New TableLayoutPanel()
            root.Dock = DockStyle.Fill
            root.Padding = New Padding(10)
            root.ColumnCount = 1
            root.RowCount = 2
            root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

            Dim top As New TableLayoutPanel()
            top.Dock = DockStyle.Top
            top.ColumnCount = 1
            top.RowCount = 2
            top.AutoSize = True
            top.AutoSizeMode = AutoSizeMode.GrowAndShrink

            Dim modelRow As New TableLayoutPanel()
            modelRow.AutoSize = True
            modelRow.AutoSizeMode = AutoSizeMode.GrowAndShrink
            modelRow.ColumnCount = 2
            modelRow.RowCount = 1
            modelRow.Dock = DockStyle.Top
            modelRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            modelRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))

            _modelLabel = New Label()
            _modelLabel.Text = "Model name:"
            _modelLabel.AutoSize = True
            _modelLabel.Anchor = AnchorStyles.Left

            _modelNameTextBox = New TextBox()
            _modelNameTextBox.Width = 260
            _modelNameTextBox.Anchor = AnchorStyles.Left Or AnchorStyles.Right

            _connectButton = New Button()
            _connectButton.Text = "Connect"
            _connectButton.AutoSize = True
            _connectButton.Margin = New Padding(2)
            AddHandler _connectButton.Click, AddressOf ConnectClickedAsync

            _retrieveButton = New Button()
            _retrieveButton.Text = "Retrieve"
            _retrieveButton.AutoSize = True
            _retrieveButton.Margin = New Padding(2)
            AddHandler _retrieveButton.Click, AddressOf RetrieveClickedAsync

            _regenerateButton = New Button()
            _regenerateButton.Text = "Regenerate"
            _regenerateButton.AutoSize = True
            _regenerateButton.Margin = New Padding(2)
            AddHandler _regenerateButton.Click, AddressOf RegenerateClickedAsync

            _loadParametersButton = New Button()
            _loadParametersButton.Text = "Load Params"
            _loadParametersButton.AutoSize = True
            _loadParametersButton.Margin = New Padding(2)
            AddHandler _loadParametersButton.Click, AddressOf LoadParametersClickedAsync

            _loadFeaturesButton = New Button()
            _loadFeaturesButton.Text = "Load Features"
            AddHandler _loadFeaturesButton.Click, AddressOf LoadFeaturesClickedAsync

            _loadComponentsButton = New Button()
            _loadComponentsButton.Text = "Load Components"
            AddHandler _loadComponentsButton.Click, AddressOf LoadComponentsClickedAsync

            _buildBomButton = New Button()
            _buildBomButton.Text = "Build BOM"
            _buildBomButton.AutoSize = True
            _buildBomButton.Margin = New Padding(2)
            AddHandler _buildBomButton.Click, AddressOf BuildBomClickedAsync

            _exportBomButton = New Button()
            _exportBomButton.Text = "Export CSV"
            _exportBomButton.AutoSize = True
            _exportBomButton.Margin = New Padding(2)
            AddHandler _exportBomButton.Click, AddressOf ExportBomClicked

            _expandAllBomButton = New Button()
            _expandAllBomButton.Text = "Expand"
            _expandAllBomButton.AutoSize = True
            _expandAllBomButton.Margin = New Padding(2)
            AddHandler _expandAllBomButton.Click, AddressOf ExpandAllBomClicked

            _collapseAllBomButton = New Button()
            _collapseAllBomButton.Text = "Collapse"
            _collapseAllBomButton.AutoSize = True
            _collapseAllBomButton.Margin = New Padding(2)
            AddHandler _collapseAllBomButton.Click, AddressOf CollapseAllBomClicked

            _showItemParametersButton = New Button()
            _showItemParametersButton.Text = "Show Params"
            _showItemParametersButton.Visible = False
            _showItemParametersButton.Enabled = False
            _showItemParametersButton.AutoSize = True
            _showItemParametersButton.Margin = New Padding(2)
            AddHandler _showItemParametersButton.Click, AddressOf ShowItemParametersClickedAsync

            _setParameterButton = New Button()
            _setParameterButton.Text = "Set"
            _setParameterButton.AutoSize = True
            _setParameterButton.Margin = New Padding(2)
            AddHandler _setParameterButton.Click, AddressOf SetSelectedParameterClickedAsync

            _clearLogButton = New Button()
            _clearLogButton.Text = "Clear Log"
            _clearLogButton.AutoSize = True
            _clearLogButton.Margin = New Padding(2)
            AddHandler _clearLogButton.Click, Sub(sender, args) _outputTextBox.Clear()

            modelRow.Controls.Add(_modelLabel, 0, 0)
            modelRow.Controls.Add(_modelNameTextBox, 1, 0)

            Dim buttonGrid As New TableLayoutPanel()
            buttonGrid.AutoSize = True
            buttonGrid.AutoSizeMode = AutoSizeMode.GrowAndShrink
            buttonGrid.ColumnCount = 5
            buttonGrid.RowCount = 1
            buttonGrid.Dock = DockStyle.Top
            For index = 0 To 4
                buttonGrid.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            Next

            Dim topButtons As Button() = {
                _connectButton,
                _retrieveButton,
                _regenerateButton,
                _loadParametersButton,
                _buildBomButton,
                _exportBomButton,
                _expandAllBomButton,
                _collapseAllBomButton,
                _showItemParametersButton,
                _setParameterButton,
                _clearLogButton
            }

            For index = 0 To topButtons.Length - 1
                Dim rowIndex = index \ 5
                Dim columnIndex = index Mod 5
                While buttonGrid.RowCount <= rowIndex
                    buttonGrid.RowCount += 1
                    buttonGrid.RowStyles.Add(New RowStyle(SizeType.AutoSize))
                End While

                buttonGrid.Controls.Add(topButtons(index), columnIndex, rowIndex)
            Next

            top.Controls.Add(modelRow, 0, 0)
            top.Controls.Add(buttonGrid, 0, 1)

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
            _paramsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
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
            _paramsGrid.Columns("Name").MinimumWidth = 220
            _paramsGrid.Columns("Name").Width = 220
            _paramsGrid.Columns("Type").Width = 80
            _paramsGrid.Columns("Value").Width = Math.Min(420, TextRenderer.MeasureText(New String("W"c, 60), _paramsGrid.Font).Width + 18)
            _paramsGrid.Columns("NewValue").Width = 140
            _paramsGrid.Columns("Description").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            _paramsGrid.Columns("Description").MinimumWidth = 220

            _featuresGrid = New DataGridView()
            _featuresGrid.Dock = DockStyle.Fill
            _featuresGrid.AllowUserToAddRows = False
            _featuresGrid.AllowUserToDeleteRows = False
            _featuresGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect 
            _featuresGrid.MultiSelect = False
            _featuresGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            _featuresGrid.ReadOnly = True
            _featuresGrid.Columns.Add("Id", "Feature ID")
            _featuresGrid.Columns.Add("Name", "Name")
            _featuresGrid.Columns.Add("TypeName", "Type")
            _featuresGrid.Columns.Add("SubTypeName", "Sub Type")
            _featuresGrid.Columns.Add("Status", "Status")

            _componentsGrid = New DataGridView()
            _componentsGrid.Dock = DockStyle.Fill
            _componentsGrid.AllowUserToAddRows = False
            _componentsGrid.AllowUserToDeleteRows = False
            _componentsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            _componentsGrid.MultiSelect = False
            _componentsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            _componentsGrid.ReadOnly = True
            _componentsGrid.Columns.Add("FeatureId", "Feature ID")
            _componentsGrid.Columns.Add("ComponentName", "Component Name")
            _componentsGrid.Columns.Add("FileName", "File Name")

            _bomGrid = New DataGridView()
            _bomGrid.Dock = DockStyle.Fill
            _bomGrid.AllowUserToAddRows = False
            _bomGrid.AllowUserToDeleteRows = False
            _bomGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            _bomGrid.MultiSelect = False
            _bomGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            _bomGrid.ReadOnly = True
            Dim toggleColumn As New DataGridViewButtonColumn()
            toggleColumn.Name = "Toggle"
            toggleColumn.HeaderText = ""
            toggleColumn.Width = 34
            toggleColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            _bomGrid.Columns.Add(toggleColumn)
            Dim levelColumn As New DataGridViewTextBoxColumn()
            levelColumn.Name = "Level"
            levelColumn.HeaderText = "Level"
            levelColumn.Width = 56
            levelColumn.MinimumWidth = 56
            levelColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            _bomGrid.Columns.Add(levelColumn)
            Dim displayNameColumn As New DataGridViewTextBoxColumn()
            displayNameColumn.Name = "DisplayName"
            displayNameColumn.HeaderText = "Name"
            displayNameColumn.Width = 240
            _bomGrid.Columns.Add(displayNameColumn)
            Dim fileNameColumn As New DataGridViewTextBoxColumn()
            fileNameColumn.Name = "FileName"
            fileNameColumn.HeaderText = "File Name"
            fileNameColumn.Width = 180
            _bomGrid.Columns.Add(fileNameColumn)
            Dim descpColumn As New DataGridViewTextBoxColumn()
            descpColumn.Name = "Descp"
            descpColumn.HeaderText = "Descp"
            descpColumn.Width = 180
            _bomGrid.Columns.Add(descpColumn)
            Dim bsClassColumn As New DataGridViewTextBoxColumn()
            bsClassColumn.Name = "BsClass"
            bsClassColumn.HeaderText = "BS_Class"
            bsClassColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            bsClassColumn.MinimumWidth = 140
            _bomGrid.Columns.Add(bsClassColumn)
            AddHandler _bomGrid.CellContentClick, AddressOf BomGridCellContentClick
            AddHandler _bomGrid.SelectionChanged, AddressOf BomGridSelectionChanged

            _itemParamsPanel = New Panel()
            _itemParamsPanel.Dock = DockStyle.Right
            _itemParamsPanel.Width = ItemPanelMinWidth
            _itemParamsPanel.MinimumSize = New Size(ItemPanelMinWidth, 0)
            _itemParamsPanel.Padding = New Padding(8)
            _itemParamsPanel.BackColor = Color.FromArgb(245, 245, 245)

            Dim itemPanelLayout As New TableLayoutPanel()
            itemPanelLayout.Dock = DockStyle.Fill
            itemPanelLayout.RowCount = 4
            itemPanelLayout.ColumnCount = 1
            itemPanelLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            itemPanelLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            itemPanelLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            itemPanelLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            _itemParamsTitle = New Label()
            _itemParamsTitle.AutoSize = True
            _itemParamsTitle.Text = "Item Parameters"
            _itemParamsTitle.Font = New Font(Font, FontStyle.Bold)

            _itemParamsLoadingBar = New ProgressBar()
            _itemParamsLoadingBar.Dock = DockStyle.Top
            _itemParamsLoadingBar.Style = ProgressBarStyle.Marquee
            _itemParamsLoadingBar.MarqueeAnimationSpeed = 24
            _itemParamsLoadingBar.Visible = False

            _itemParamsGrid = New DataGridView()
            _itemParamsGrid.Dock = DockStyle.Fill
            _itemParamsGrid.AllowUserToAddRows = False
            _itemParamsGrid.AllowUserToDeleteRows = False
            _itemParamsGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            _itemParamsGrid.MultiSelect = False
            _itemParamsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            _itemParamsGrid.Columns.Add("Name", "Parameter")
            _itemParamsGrid.Columns.Add("Type", "Type")
            _itemParamsGrid.Columns.Add("Value", "Current Value")
            _itemParamsGrid.Columns.Add("NewValue", "New Value")
            _itemParamsGrid.Columns.Add("Description", "Description")
            Dim itemDiscrColumn = New DataGridViewTextBoxColumn()
            itemDiscrColumn.Name = "Discr"
            itemDiscrColumn.Visible = False
            _itemParamsGrid.Columns.Add(itemDiscrColumn)
            _itemParamsGrid.Columns("Name").ReadOnly = True
            _itemParamsGrid.Columns("Type").ReadOnly = True
            _itemParamsGrid.Columns("Value").ReadOnly = True
            _itemParamsGrid.Columns("Description").ReadOnly = True
            _itemParamsGrid.Columns("Name").MinimumWidth = 220
            _itemParamsGrid.Columns("Name").Width = 220
            _itemParamsGrid.Columns("Type").Width = 80
            _itemParamsGrid.Columns("Value").Width = Math.Min(420, TextRenderer.MeasureText(New String("W"c, 60), _itemParamsGrid.Font).Width + 18)
            _itemParamsGrid.Columns("NewValue").Width = 140
            _itemParamsGrid.Columns("Description").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            _itemParamsGrid.Columns("Description").MinimumWidth = 220

            Dim itemPanelButtons As New FlowLayoutPanel()
            itemPanelButtons.Dock = DockStyle.Fill
            itemPanelButtons.FlowDirection = FlowDirection.RightToLeft
            itemPanelButtons.AutoSize = True

            _itemSetParameterButton = New Button()
            _itemSetParameterButton.Text = "Set Selected + Regenerate"
            AddHandler _itemSetParameterButton.Click, AddressOf SetSelectedItemParameterClickedAsync

            itemPanelButtons.Controls.Add(_itemSetParameterButton)

            itemPanelLayout.Controls.Add(_itemParamsTitle, 0, 0)
            itemPanelLayout.Controls.Add(_itemParamsLoadingBar, 0, 1)
            itemPanelLayout.Controls.Add(_itemParamsGrid, 0, 2)
            itemPanelLayout.Controls.Add(itemPanelButtons, 0, 3)
            _itemParamsPanel.Controls.Add(itemPanelLayout)

            _tabs = New TabControl()
            _tabs.Dock = DockStyle.Fill

            Dim logTab As New TabPage("Log")
            logTab.Controls.Add(_outputTextBox)

            Dim paramsTab As New TabPage("Parameters")
            paramsTab.Controls.Add(_paramsGrid)

            Dim bomTab As New TabPage("BOM")
            bomTab.Controls.Add(_bomGrid)

            _tabs.TabPages.Add(logTab)
            _tabs.TabPages.Add(paramsTab)
            _tabs.TabPages.Add(bomTab)
            AddHandler _tabs.SelectedIndexChanged, AddressOf TabsSelectedIndexChanged
            AddHandler Me.Resize, AddressOf MainFormResized

            EnableDoubleBuffering(_paramsGrid)
            EnableDoubleBuffering(_featuresGrid)
            EnableDoubleBuffering(_componentsGrid)
            EnableDoubleBuffering(_bomGrid)
            EnableDoubleBuffering(_itemParamsGrid)

            root.Controls.Add(top, 0, 0)
            root.Controls.Add(_tabs, 0, 1)

            Controls.Add(root)
            Controls.Add(_itemParamsPanel)

            ApplyDefaultLayoutSettings()
            UpdateItemPanelWidth()
            UpdateItemPanelVisibility()

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

        Private Async Sub LoadFeaturesClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            SetBusyState(True)
            Try
                AddLog("Load Features started...")
                Dim result = Await _service.GetFeaturesAsync(modelInput)
                If Not result.Success Then
                    AddLog("ERROR: " & result.Message)
                    Return
                End If

                _featuresGrid.Rows.Clear()
                For Each item In result.Items
                    _featuresGrid.Rows.Add(item.Id, item.Name, item.TypeName, item.SubTypeName, item.Status)
                Next

                _tabs.SelectedIndex = 2 ' Activate Features tab
                AddLog("OK: " & result.Message)
            Catch ex As Exception
                AddLog("ERROR: " & ex.Message)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Async Sub LoadComponentsClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            SetBusyState(True)
            Try
                AddLog("Load Components started...")
                Dim result = Await _service.GetComponentsAsync(modelInput)
                If Not result.Success Then
                    AddLog("ERROR: " & result.Message)
                    Return
                End If

                _componentsGrid.Rows.Clear()
                For Each item In result.Items
                    _componentsGrid.Rows.Add(item.FeatureId, item.ComponentName, item.FileName)
                Next

                _tabs.SelectedIndex = 3
                AddLog("OK: " & result.Message)
            Catch ex As Exception
                AddLog("ERROR: " & ex.Message)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Async Sub BuildBomClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            SetBusyState(True)
            Try
                AddLog("Build BOM started...")
                Dim result = Await _service.GetBomHierarchyAsync(modelInput)
                If Not result.Success Then
                    AddLog("ERROR: " & result.Message)
                    Return
                End If

                _bomRowsAll.Clear()
                _bomRowsByNodeId.Clear()
                _bomCollapsedNodeIds.Clear()
                For Each row In result.Rows
                    _bomRowsAll.Add(row)
                    _bomRowsByNodeId(row.NodeId) = row
                Next

                RenderBomGrid()
                UpdateShowParametersButtonState()
                WindowState = FormWindowState.Maximized

                _tabs.SelectedIndex = 2
                AddLog("OK: " & result.Message)
            Catch ex As Exception
                AddLog("ERROR: " & ex.Message)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Async Sub SetSelectedParameterClickedAsync(sender As Object, e As EventArgs)
            Dim updates = CollectPendingParameterUpdates(_paramsGrid)
            If updates.Count = 0 Then
                AddLog("ERROR: Enter at least one value in the 'New Value' column before setting.")
                Return
            End If

            Dim modelInput = _modelNameTextBox.Text
            SetBusyState(True)
            Try
                AddLog("Set Parameters started...")
                Dim result = Await _service.SetParametersAsync(modelInput, updates)
                If Not result.Success Then
                    AddLog("ERROR: " & result.Message)
                    Return
                End If

                ClearPendingParameterUpdates(_paramsGrid)
                Await ReloadParametersGridAsync(_paramsGrid, modelInput)
                _tabs.SelectedIndex = 1
                AddLog("OK: " & result.Message)
            Catch ex As Exception
                AddLog("ERROR: " & ex.Message)
            Finally
                SetBusyState(False)
            End Try
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

        Private Function CreateIndentedLevel(level As Integer) As String
            If level <= 0 Then
                Return level.ToString()
            End If

            Return New String(" "c, level * 4) & level.ToString()
        End Function

        Private Sub BomGridCellContentClick(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex < 0 OrElse e.ColumnIndex <> _bomGrid.Columns("Toggle").Index Then
                Return
            End If

            Dim nodeIdObj = _bomGrid.Rows(e.RowIndex).Tag
            If nodeIdObj Is Nothing Then
                Return
            End If

            Dim nodeId = CInt(nodeIdObj)
            If Not _bomRowsByNodeId.ContainsKey(nodeId) Then
                Return
            End If

            Dim row = _bomRowsByNodeId(nodeId)
            If Not row.HasChildren Then
                Return
            End If

            If _bomCollapsedNodeIds.Contains(nodeId) Then
                _bomCollapsedNodeIds.Remove(nodeId)
            Else
                _bomCollapsedNodeIds.Add(nodeId)
            End If

            RenderBomGrid()
        End Sub

        Private Sub RenderBomGrid()
            _suppressSelectionChanged = True
            Try
                _bomGrid.Rows.Clear()

                For Each row In _bomRowsAll
                    If Not IsRowVisible(row) Then
                        Continue For
                    End If

                    Dim toggleText = ""
                    If row.HasChildren Then
                        toggleText = If(_bomCollapsedNodeIds.Contains(row.NodeId), "+", "-")
                    End If

                    Dim rowIndex = _bomGrid.Rows.Add(toggleText, CreateIndentedLevel(row.Level), row.DisplayName, row.FileName, row.Descp, row.BsClass)
                    _bomGrid.Rows(rowIndex).Tag = row.NodeId
                Next
            Finally
                _suppressSelectionChanged = False
            End Try

            UpdateShowParametersButtonState()
        End Sub

        Private Async Sub BomGridSelectionChanged(sender As Object, e As EventArgs)
            UpdateShowParametersButtonState()
            If _suppressSelectionChanged Then Return
            Await LoadItemParametersForSelectedRowAsync(forceReload:=False)
        End Sub

        Private Sub TabsSelectedIndexChanged(sender As Object, e As EventArgs)
            UpdateItemPanelWidth()
            UpdateItemPanelVisibility()
            UpdateShowParametersButtonState()
        End Sub

        Private Sub UpdateShowParametersButtonState()
            Dim hasSelection = _tabs.SelectedIndex = BomTabIndex AndAlso _bomGrid.SelectedRows.Count > 0
            _showItemParametersButton.Visible = hasSelection
            _showItemParametersButton.Enabled = hasSelection
        End Sub

        Private Sub UpdateItemPanelVisibility()
            Dim isBomTab = (_tabs.SelectedIndex = BomTabIndex)
            _itemParamsPanel.Visible = isBomTab
        End Sub

        Private Sub MainFormResized(sender As Object, e As EventArgs)
            UpdateItemPanelWidth()
        End Sub

        Private Sub UpdateItemPanelWidth()
            _itemParamsPanel.Width = Math.Max(ItemPanelMinWidth, Me.ClientSize.Width \ 2)
        End Sub

        Private Sub ExpandAllBomClicked(sender As Object, e As EventArgs)
            If _bomRowsAll.Count = 0 Then
                Return
            End If

            _bomCollapsedNodeIds.Clear()
            RenderBomGrid()
        End Sub

        Private Sub CollapseAllBomClicked(sender As Object, e As EventArgs)
            If _bomRowsAll.Count = 0 Then
                Return
            End If

            _bomCollapsedNodeIds.Clear()
            For Each row In _bomRowsAll
                If row.HasChildren Then
                    _bomCollapsedNodeIds.Add(row.NodeId)
                End If
            Next

            _bomCollapsedNodeIds.Remove(1)
            RenderBomGrid()
        End Sub

        Private Async Sub ShowItemParametersClickedAsync(sender As Object, e As EventArgs)
            Await LoadItemParametersForSelectedRowAsync(forceReload:=True)
        End Sub

        Private Async Function LoadItemParametersForSelectedRowAsync(forceReload As Boolean) As Task
            If _bomGrid.SelectedRows.Count = 0 Then
                Return
            End If

            If _isLoadingItemParams Then
                Return
            End If

            Dim selectedRow = _bomGrid.SelectedRows(0)
            Dim modelInput = Convert.ToString(selectedRow.Cells("FileName").Value)
            If String.IsNullOrWhiteSpace(modelInput) Then
                modelInput = Convert.ToString(selectedRow.Cells("DisplayName").Value)
            End If

            If String.IsNullOrWhiteSpace(modelInput) Then
                Return
            End If

            If Not forceReload AndAlso String.Equals(_selectedBomModelInput, modelInput, StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            _selectedBomModelInput = modelInput
            _itemParamsTitle.Text = "Item Parameters - " & modelInput

            _isLoadingItemParams = True
            SetBusyState(True)
            BeginItemParametersLoading(modelInput)
            Try
                Dim result = Await _service.GetParametersAsync(modelInput)
                If Not result.Success Then
                    AddLog("ERROR: " & result.Message)
                    Return
                End If

                _itemParamsGrid.Rows.Clear()
                For Each item In result.Items
                    Dim typeName = GetDiscrName(item.Discr)
                    _itemParamsGrid.Rows.Add(item.Name, typeName, item.ValueText, "", item.Description, item.Discr)
                Next

                AddLog("OK: Loaded item parameters for " & modelInput)
            Catch ex As Exception
                AddLog("ERROR: " & ex.Message)
            Finally
                EndItemParametersLoading(modelInput)
                SetBusyState(False)
                _isLoadingItemParams = False
            End Try
        End Function

        Private Async Sub SetSelectedItemParameterClickedAsync(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(_selectedBomModelInput) Then
                AddLog("ERROR: No selected BOM item model.")
                Return
            End If

            Dim updates = CollectPendingParameterUpdates(_itemParamsGrid)
            If updates.Count = 0 Then
                AddLog("ERROR: Enter at least one value in the 'New Value' column before setting.")
                Return
            End If

            SetBusyState(True)
            Try
                AddLog("Set Item Parameters started...")
                Dim result = Await _service.SetParametersAsync(_selectedBomModelInput, updates)
                If result.Success Then
                    ClearPendingParameterUpdates(_itemParamsGrid)
                    Await ReloadParametersGridAsync(_itemParamsGrid, _selectedBomModelInput)
                    AddLog("OK: " & result.Message)
                Else
                    AddLog("ERROR: " & result.Message)
                End If
            Catch ex As Exception
                AddLog("ERROR: " & ex.Message)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Function IsRowVisible(row As BomGridRowItem) As Boolean
            Dim parentId = row.ParentNodeId
            While parentId > 0
                If _bomCollapsedNodeIds.Contains(parentId) Then
                    Return False
                End If

                If Not _bomRowsByNodeId.ContainsKey(parentId) Then
                    Exit While
                End If

                parentId = _bomRowsByNodeId(parentId).ParentNodeId
            End While

            Return True
        End Function

        Private Sub ExportBomClicked(sender As Object, e As EventArgs)
            If _bomGrid.Rows.Count = 0 Then
                AddLog("ERROR: Build the BOM before exporting CSV.")
                Return
            End If

            Using dialog As New SaveFileDialog()
                dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
                dialog.Title = "Export BOM CSV"
                dialog.FileName = "bom-hierarchy.csv"

                If dialog.ShowDialog(Me) <> DialogResult.OK Then
                    Return
                End If

                Try
                    Using writer As New StreamWriter(dialog.FileName, False, Encoding.UTF8)
                        Dim headers As String() = {"Level", "Name", "File Name", "Descp", "BS_Class"}
                        writer.WriteLine(String.Join(",", EscapeCsvValues(headers)))

                        For Each row As DataGridViewRow In _bomGrid.Rows
                            If row.IsNewRow Then
                                Continue For
                            End If

                            Dim values As String() = {
                                Convert.ToString(row.Cells("Level").Value),
                                Convert.ToString(row.Cells("DisplayName").Value),
                                Convert.ToString(row.Cells("FileName").Value),
                                Convert.ToString(row.Cells("Descp").Value),
                                Convert.ToString(row.Cells("BsClass").Value)
                            }
                            writer.WriteLine(String.Join(",", EscapeCsvValues(values)))
                        Next
                    End Using

                    AddLog("OK: BOM CSV exported.")
                Catch ex As Exception
                    AddLog("ERROR: CSV export failed: " & ex.Message)
                End Try
            End Using
        End Sub

        Private Function EscapeCsvValues(values As IEnumerable(Of String)) As IEnumerable(Of String)
            Dim escaped As New List(Of String)()
            For Each value In values
                Dim text = If(value, String.Empty)
                text = text.Replace("""", """""")
                escaped.Add("""" & text & """")
            Next
            Return escaped
        End Function

        Private Sub SetBusyState(isBusy As Boolean)
            UseWaitCursor = isBusy
            _connectButton.Enabled = Not isBusy
            _retrieveButton.Enabled = Not isBusy
            _regenerateButton.Enabled = Not isBusy
            _loadParametersButton.Enabled = Not isBusy
            _loadFeaturesButton.Enabled = Not isBusy
            _loadComponentsButton.Enabled = Not isBusy
            _buildBomButton.Enabled = Not isBusy
            _exportBomButton.Enabled = Not isBusy
            _expandAllBomButton.Enabled = Not isBusy
            _collapseAllBomButton.Enabled = Not isBusy
            _showItemParametersButton.Enabled = Not isBusy AndAlso _tabs.SelectedIndex = BomTabIndex AndAlso _bomGrid.SelectedRows.Count > 0
            _setParameterButton.Enabled = Not isBusy
            _clearLogButton.Enabled = Not isBusy
            _itemSetParameterButton.Enabled = Not isBusy
            _itemParamsGrid.Enabled = Not isBusy
        End Sub

        Private Sub BeginItemParametersLoading(modelInput As String)
            _itemParamsLoadingBar.Visible = True
            _itemParamsTitle.Text = "Item Parameters - " & modelInput & " (Loading...)"
        End Sub

        Private Sub EndItemParametersLoading(modelInput As String)
            _itemParamsLoadingBar.Visible = False
            _itemParamsTitle.Text = "Item Parameters - " & modelInput
        End Sub

        Private Shared Sub EnableDoubleBuffering(control As Control)
            Dim propertyInfo = control.GetType().GetProperty("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance)
            If propertyInfo IsNot Nothing Then
                propertyInfo.SetValue(control, True, Nothing)
            End If
        End Sub

        Private Function CollectPendingParameterUpdates(grid As DataGridView) As List(Of ParameterUpdateItem)
            Dim updates As New List(Of ParameterUpdateItem)()

            For Each row As DataGridViewRow In grid.Rows
                If row.IsNewRow Then
                    Continue For
                End If

                Dim parameterName = Convert.ToString(row.Cells("Name").Value)
                Dim userInput = Convert.ToString(row.Cells("NewValue").Value)
                If String.IsNullOrWhiteSpace(parameterName) OrElse String.IsNullOrWhiteSpace(userInput) Then
                    Continue For
                End If

                updates.Add(New ParameterUpdateItem With {
                    .Name = parameterName,
                    .UserInput = userInput
                })
            Next

            Return updates
        End Function

        Private Sub ClearPendingParameterUpdates(grid As DataGridView)
            For Each row As DataGridViewRow In grid.Rows
                If row.IsNewRow Then
                    Continue For
                End If

                row.Cells("NewValue").Value = ""
            Next
        End Sub

        Private Async Function ReloadParametersGridAsync(grid As DataGridView, modelInput As String) As Task
            Dim result = Await _service.GetParametersAsync(modelInput)
            If Not result.Success Then
                Throw New InvalidOperationException(result.Message)
            End If

            grid.Rows.Clear()
            For Each item In result.Items
                Dim typeName = GetDiscrName(item.Discr)
                grid.Rows.Add(item.Name, typeName, item.ValueText, "", item.Description, item.Discr)
            Next
        End Function

        Private Sub ApplyDefaultLayoutSettings()
            Width = 760
            Height = 480
            _itemParamsPanel.Width = Math.Max(ItemPanelMinWidth, Width \ 2)

            _paramsGrid.Columns("Name").Width = 220
            _paramsGrid.Columns("Type").Width = 80
            _paramsGrid.Columns("Value").Width = Math.Min(420, TextRenderer.MeasureText(New String("W"c, 60), _paramsGrid.Font).Width + 18)
            _paramsGrid.Columns("NewValue").Width = 140

            _itemParamsGrid.Columns("Name").Width = 220
            _itemParamsGrid.Columns("Type").Width = 80
            _itemParamsGrid.Columns("Value").Width = Math.Min(420, TextRenderer.MeasureText(New String("W"c, 60), _itemParamsGrid.Font).Width + 18)
            _itemParamsGrid.Columns("NewValue").Width = 140

            _bomGrid.Columns("Toggle").Width = 34
            _bomGrid.Columns("Level").Width = 56
            _bomGrid.Columns("DisplayName").Width = 240
            _bomGrid.Columns("FileName").Width = 180
            _bomGrid.Columns("Descp").Width = 180
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
