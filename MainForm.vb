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
        Private Const ToolbarGroupGap As Integer = 16
        Private Const MaxCsvFilenameLength As Integer = 255
        Private Const GenericUiErrorMessage As String = "The operation could not be completed."

        Private ReadOnly _service As CreoSessionService

        Private ReadOnly _modelLabel As Label
        Private ReadOnly _modelNameTextBox As TextBox
        Private ReadOnly _connectButton As Button
        Private ReadOnly _retrieveButton As Button
        Private ReadOnly _regenerateButton As Button
        Private ReadOnly _exportFormatComboBox As ComboBox
        Private ReadOnly _exportModelButton As Button
        Private ReadOnly _loadParametersButton As Button
        Private ReadOnly _loadFeaturesButton As Button
        Private ReadOnly _loadComponentsButton As Button
        Private ReadOnly _buildBomButton As Button
        Private ReadOnly _exportBomButton As Button
        Private ReadOnly _expandAllBomButton As Button
        Private ReadOnly _collapseAllBomButton As Button
        Private ReadOnly _applyConfigButton As Button
        Private ReadOnly _saveRulesButton As Button
        Private ReadOnly _pushToModelsButton As Button
        Private ReadOnly _hideExcludedButton As Button
        Private ReadOnly _showItemParametersButton As Button
        Private ReadOnly _setParameterButton As Button
        Private ReadOnly _clearLogButton As Button
        Private ReadOnly _activateButton As Button
        Private ReadOnly _outputTextBox As TextBox
        Private ReadOnly _tabs As TabControl
        Private ReadOnly _paramsGrid As DataGridView
        Private ReadOnly _featuresGrid As DataGridView
        Private ReadOnly _componentsGrid As DataGridView
        Private ReadOnly _bomGrid As DataGridView
        Private ReadOnly _bomRowsAll As List(Of BomGridRowItem)
        Private ReadOnly _bomRowsByNodeId As Dictionary(Of Integer, BomGridRowItem)
        Private ReadOnly _bomCollapsedNodeIds As HashSet(Of Integer)
        Private ReadOnly _allowedExportDirectories As List(Of String)
        Private ReadOnly _itemParamsPanel As Panel
        Private ReadOnly _itemParamsTitle As Label
        Private ReadOnly _itemParamsLoadingBar As ProgressBar
        Private ReadOnly _itemParamsGrid As DataGridView
        Private ReadOnly _itemSetParameterButton As Button
        Private ReadOnly _itemParamsCloseButton As Button
        Private _selectedBomModelInput As String
        Private _suppressSelectionChanged As Boolean
        Private _isLoadingItemParams As Boolean
        Private _licenseStatus As LicenseStatus
        Private _hideExcluded As Boolean
        Private ReadOnly _strikeoutFont As Font
        Private _paramsModelInput As String
        Private _bomModelInput As String
        Private _bomTopFileName As String
        Private _bomRulesDirty As Boolean

        Friend Sub New(status As LicenseStatus)
            _licenseStatus = status
            Text = BuildWindowTitle(status)
            Width = 760
            Height = 480
            MinimumSize = New Size(560, 360)
            StartPosition = FormStartPosition.CenterScreen
            Font = New Font("Segoe UI", 9.0F)
            SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.AllPaintingInWmPaint, True)
            UpdateStyles()

            _service = New CreoSessionService()
            _bomRowsAll = New List(Of BomGridRowItem)()
            _bomRowsByNodeId = New Dictionary(Of Integer, BomGridRowItem)()
            _bomCollapsedNodeIds = New HashSet(Of Integer)()
            _allowedExportDirectories = New List(Of String)()
            InitializeExportDirectories()
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

            _connectButton = CreateToolbarButton("Connect", AddressOf ConnectClickedAsync)
            _retrieveButton = CreateToolbarButton("Retrieve", AddressOf RetrieveClickedAsync)
            _regenerateButton = CreateToolbarButton("Regenerate", AddressOf RegenerateClickedAsync)

            _exportFormatComboBox = New ComboBox()
            _exportFormatComboBox.DropDownStyle = ComboBoxStyle.DropDownList
            _exportFormatComboBox.Items.AddRange(New Object() {"PDF", "DXF", "STEP", "IGES"})
            _exportFormatComboBox.SelectedIndex = 0
            _exportFormatComboBox.Width = 80
            _exportFormatComboBox.Margin = New Padding(ToolbarGroupGap, 5, 3, 3)

            _exportModelButton = CreateToolbarButton("Export Model", AddressOf ExportModelClickedAsync)
            _loadParametersButton = CreateToolbarButton("Load Params", AddressOf LoadParametersClickedAsync)

            _loadFeaturesButton = New Button()
            _loadFeaturesButton.Text = "Load Features"
            AddHandler _loadFeaturesButton.Click, AddressOf LoadFeaturesClickedAsync

            _loadComponentsButton = New Button()
            _loadComponentsButton.Text = "Load Components"
            AddHandler _loadComponentsButton.Click, AddressOf LoadComponentsClickedAsync

            _buildBomButton = CreateToolbarButton("Build BOM", AddressOf BuildBomClickedAsync)
            _expandAllBomButton = CreateToolbarButton("Expand", AddressOf ExpandAllBomClicked)
            _collapseAllBomButton = CreateToolbarButton("Collapse", AddressOf CollapseAllBomClicked)
            _hideExcludedButton = CreateToolbarButton("Hide Excluded", AddressOf HideExcludedClicked)

            _applyConfigButton = CreateToolbarButton("Apply Config", AddressOf ApplyConfigClicked)
            _applyConfigButton.Margin = New Padding(ToolbarGroupGap, 3, 3, 3)
            _saveRulesButton = CreateToolbarButton("Save Rules", AddressOf SaveRulesClickedAsync)
            _pushToModelsButton = CreateToolbarButton("Push to Models", AddressOf PushToModelsClickedAsync)

            _exportBomButton = CreateToolbarButton("Export CSV", AddressOf ExportBomClicked)
            _exportBomButton.Margin = New Padding(ToolbarGroupGap, 3, 3, 3)
            _showItemParametersButton = CreateToolbarButton("Show Params", AddressOf ShowItemParametersClickedAsync)
            _showItemParametersButton.Visible = False
            _showItemParametersButton.Enabled = False

            _setParameterButton = CreateToolbarButton("Set", AddressOf SetSelectedParameterClickedAsync)
            _clearLogButton = CreateToolbarButton("Clear Log", Sub(sender, args) _outputTextBox.Clear())

            _activateButton = CreateToolbarButton("Activate", AddressOf ActivateClicked)
            _activateButton.Margin = New Padding(ToolbarGroupGap, 3, 3, 3)
            _activateButton.Visible = status.State <> LicenseState.Licensed

            modelRow.Controls.Add(_modelLabel, 0, 0)
            modelRow.Controls.Add(_modelNameTextBox, 1, 0)

            Dim globalBar As New FlowLayoutPanel()
            globalBar.Dock = DockStyle.Fill
            globalBar.AutoSize = True
            globalBar.AutoSizeMode = AutoSizeMode.GrowAndShrink
            globalBar.WrapContents = True
            globalBar.Margin = New Padding(0, 6, 0, 0)
            globalBar.Controls.AddRange(New Control() {
                _connectButton,
                _retrieveButton,
                _regenerateButton,
                _exportFormatComboBox,
                _exportModelButton,
                _activateButton
            })

            top.Controls.Add(modelRow, 0, 0)
            top.Controls.Add(globalBar, 0, 1)

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
            _bomGrid.ReadOnly = False
            _bomGrid.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
            _bomGrid.ShowCellToolTips = True
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
            Dim featureIdColumn As New DataGridViewTextBoxColumn()
            featureIdColumn.Name = "FeatureId"
            featureIdColumn.HeaderText = "Feature ID"
            featureIdColumn.Width = 70
            featureIdColumn.MinimumWidth = 56
            featureIdColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            _bomGrid.Columns.Add(featureIdColumn)
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
            bsClassColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            bsClassColumn.Width = 140
            bsClassColumn.MinimumWidth = 100
            _bomGrid.Columns.Add(bsClassColumn)
            Dim conditionsColumn As New DataGridViewTextBoxColumn()
            conditionsColumn.Name = "Conditions"
            conditionsColumn.HeaderText = "Conditions"
            conditionsColumn.Width = 220
            conditionsColumn.MinimumWidth = 80
            conditionsColumn.MaxInputLength = BomRuleLimits.MaxRuleLength
            conditionsColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            _bomGrid.Columns.Add(conditionsColumn)
            Dim inheritanceColumn As New DataGridViewTextBoxColumn()
            inheritanceColumn.Name = "Inheritance"
            inheritanceColumn.HeaderText = "Inheritance"
            inheritanceColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            inheritanceColumn.MinimumWidth = 80
            inheritanceColumn.MaxInputLength = BomRuleLimits.MaxRuleLength
            _bomGrid.Columns.Add(inheritanceColumn)
            Dim resolvedColumn As New DataGridViewTextBoxColumn()
            resolvedColumn.Name = "Resolved"
            resolvedColumn.HeaderText = "Resolved Values"
            resolvedColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            resolvedColumn.MinimumWidth = 80
            _bomGrid.Columns.Add(resolvedColumn)
            For Each column As DataGridViewColumn In _bomGrid.Columns
                column.ReadOnly = column IsNot conditionsColumn AndAlso column IsNot inheritanceColumn
            Next
            _strikeoutFont = New Font(_bomGrid.Font, FontStyle.Strikeout)
            AddHandler _bomGrid.CellEndEdit, AddressOf BomGridCellEndEdit
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

            Dim titleRow As New TableLayoutPanel()
            titleRow.Dock = DockStyle.Fill
            titleRow.AutoSize = True
            titleRow.AutoSizeMode = AutoSizeMode.GrowAndShrink
            titleRow.ColumnCount = 2
            titleRow.RowCount = 1
            titleRow.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            titleRow.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            _itemParamsTitle = New Label()
            _itemParamsTitle.AutoSize = True
            _itemParamsTitle.Text = "Item Parameters"
            _itemParamsTitle.Font = New Font(Font, FontStyle.Bold)
            _itemParamsTitle.Anchor = AnchorStyles.Left

            _itemParamsCloseButton = New Button()
            _itemParamsCloseButton.Text = ChrW(&H2716)
            _itemParamsCloseButton.Size = New Size(24, 24)
            _itemParamsCloseButton.FlatStyle = FlatStyle.Flat
            _itemParamsCloseButton.FlatAppearance.BorderSize = 0
            _itemParamsCloseButton.Cursor = Cursors.Hand
            _itemParamsCloseButton.Anchor = AnchorStyles.Right
            AddHandler _itemParamsCloseButton.Click, AddressOf CloseItemParametersPanel

            titleRow.Controls.Add(_itemParamsTitle, 0, 0)
            titleRow.Controls.Add(_itemParamsCloseButton, 1, 0)

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

            itemPanelLayout.Controls.Add(titleRow, 0, 0)
            itemPanelLayout.Controls.Add(_itemParamsLoadingBar, 0, 1)
            itemPanelLayout.Controls.Add(_itemParamsGrid, 0, 2)
            itemPanelLayout.Controls.Add(itemPanelButtons, 0, 3)
            _itemParamsPanel.Controls.Add(itemPanelLayout)

            _tabs = New TabControl()
            _tabs.Dock = DockStyle.Fill

            Dim logTab = CreateToolbarTabPage("Log", _outputTextBox, _clearLogButton)
            Dim paramsTab = CreateToolbarTabPage("Parameters", _paramsGrid, _loadParametersButton, _setParameterButton)
            Dim bomTab = CreateToolbarTabPage("BOM", _bomGrid,
                _buildBomButton, _expandAllBomButton, _collapseAllBomButton, _hideExcludedButton,
                _applyConfigButton, _saveRulesButton, _pushToModelsButton,
                _exportBomButton, _showItemParametersButton)

            _tabs.TabPages.Add(logTab)
            _tabs.TabPages.Add(paramsTab)
            _tabs.TabPages.Add(bomTab)
            AddHandler _tabs.SelectedIndexChanged, AddressOf TabsSelectedIndexChanged
            AddHandler Me.Resize, AddressOf MainFormResized

            For Each grid In New DataGridView() {_paramsGrid, _featuresGrid, _componentsGrid, _bomGrid, _itemParamsGrid}
                EnableDoubleBuffering(grid)
                grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                grid.RowTemplate.Height = grid.Font.Height + 8
            Next

            root.Controls.Add(top, 0, 0)
            root.Controls.Add(_tabs, 0, 1)

            Controls.Add(root)
            Controls.Add(_itemParamsPanel)

            ApplyDefaultLayoutSettings()
            UpdateItemPanelWidth()
            UpdateItemPanelVisibility()

            AddLog("UI ready.")
            AddLog("Example: model or model.prt or myassy.asm")
        End Sub

        Private Shared Function CreateToolbarButton(text As String, handler As EventHandler) As Button
            Dim button As New Button()
            button.Text = text
            button.AutoSize = True
            button.MinimumSize = New Size(88, 28)
            button.Padding = New Padding(6, 0, 6, 0)
            button.Margin = New Padding(3)
            AddHandler button.Click, handler
            Return button
        End Function

        Private Shared Function CreateToolbarTabPage(title As String, content As Control, ParamArray tools As Control()) As TabPage
            Dim toolbar As New FlowLayoutPanel()
            toolbar.Dock = DockStyle.Fill
            toolbar.AutoSize = True
            toolbar.AutoSizeMode = AutoSizeMode.GrowAndShrink
            toolbar.WrapContents = True
            toolbar.Padding = New Padding(2, 4, 2, 2)
            toolbar.Controls.AddRange(tools)

            Dim pageLayout As New TableLayoutPanel()
            pageLayout.Dock = DockStyle.Fill
            pageLayout.ColumnCount = 1
            pageLayout.RowCount = 2
            pageLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            pageLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            pageLayout.Controls.Add(toolbar, 0, 0)
            pageLayout.Controls.Add(content, 0, 1)

            Dim page As New TabPage(title)
            page.Controls.Add(pageLayout)
            Return page
        End Function

        Private Sub InitializeExportDirectories()
            _allowedExportDirectories.Clear()
            AddAllowedExportDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
            AddAllowedExportDirectory(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory))

            Dim downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")
            AddAllowedExportDirectory(downloadsPath)
            AddAllowedExportDirectory(Path.GetTempPath())
        End Sub

        Private Sub AddAllowedExportDirectory(directoryPath As String)
            If String.IsNullOrWhiteSpace(directoryPath) Then
                Return
            End If

            If Not Directory.Exists(directoryPath) Then
                Return
            End If

            Dim normalizedPath = NormalizeDirectoryPath(directoryPath)
            For Each allowedPath In _allowedExportDirectories
                If String.Equals(allowedPath, normalizedPath, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If
            Next

            _allowedExportDirectories.Add(normalizedPath)
        End Sub

        Private Function NormalizeDirectoryPath(directoryPath As String) As String
            Return Path.GetFullPath(directoryPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        End Function

        Private Function ValidateExportPath(filePath As String) As Tuple(Of Boolean, String)
            If String.IsNullOrWhiteSpace(filePath) Then
                Return Tuple.Create(False, "Choose a CSV export location.")
            End If

            Dim fullPath As String
            Try
                fullPath = Path.GetFullPath(filePath.Trim())
            Catch
                Return Tuple.Create(False, "Choose a valid CSV export location.")
            End Try

            If fullPath.StartsWith("\\", StringComparison.Ordinal) Then
                Return Tuple.Create(False, "CSV export must use a local folder.")
            End If

            Dim directoryPath = Path.GetDirectoryName(fullPath)
            If String.IsNullOrWhiteSpace(directoryPath) OrElse Not Directory.Exists(directoryPath) Then
                Return Tuple.Create(False, "Choose an existing CSV export folder.")
            End If

            If Not IsAllowedExportDirectory(directoryPath) Then
                Return Tuple.Create(False, "CSV export is limited to Documents, Desktop, Downloads, or Temp.")
            End If

            Dim fileName = Path.GetFileName(fullPath)
            If String.IsNullOrWhiteSpace(fileName) Then
                Return Tuple.Create(False, "Choose a CSV file name.")
            End If

            If fileName.Length > MaxCsvFilenameLength Then
                Return Tuple.Create(False, "The CSV file name is too long.")
            End If

            If fileName.Contains(":"c) Then
                Return Tuple.Create(False, "The CSV file name is invalid.")
            End If

            If fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 Then
                Return Tuple.Create(False, "The CSV file name is invalid.")
            End If

            If Not String.Equals(Path.GetExtension(fullPath), ".csv", StringComparison.OrdinalIgnoreCase) Then
                Return Tuple.Create(False, "CSV export requires a .csv file name.")
            End If

            Return Tuple.Create(True, fullPath)
        End Function

        Private Function IsAllowedExportDirectory(directoryPath As String) As Boolean
            Dim normalizedDirectory = NormalizeDirectoryPath(directoryPath)
            For Each allowedDirectory In _allowedExportDirectories
                If String.Equals(normalizedDirectory, allowedDirectory, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If

                If normalizedDirectory.StartsWith(allowedDirectory & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Sub HandleUnexpectedError(context As String, ex As Exception, Optional userMessage As String = Nothing)
            CreoSessionService.LogApplicationError(context, ex)
            AddLog("ERROR: " & If(String.IsNullOrWhiteSpace(userMessage), GenericUiErrorMessage, userMessage))
        End Sub

        Private Function ValidateModelInput(ByRef modelInput As String) As Boolean
            If String.IsNullOrWhiteSpace(modelInput) Then
                AddLog("ERROR: Enter a model name.")
                Return False
            End If

            modelInput = modelInput.Trim()

            If modelInput.Length > 260 Then
                AddLog("ERROR: Model name is too long.")
                Return False
            End If

            If modelInput.IndexOf("\"c) >= 0 OrElse modelInput.IndexOf("/"c) >= 0 OrElse modelInput.Contains("..") Then
                AddLog("ERROR: Enter a Creo model name, not a file path.")
                Return False
            End If

            For Each ch In modelInput
                If Char.IsControl(ch) Then
                    AddLog("ERROR: Model name contains invalid characters.")
                    Return False
                End If
            Next

            Return True
        End Function

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
            If Not ValidateModelInput(modelInput) Then Return
            Await RunOperationAsync(Function() _service.RetrieveModelAsync(modelInput), "Retrieve")
        End Sub

        Private Async Sub RegenerateClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            If Not ValidateModelInput(modelInput) Then Return
            Await RunOperationAsync(Function() _service.RegenerateModelAsync(modelInput), "Regenerate")
        End Sub

        Private Async Sub ExportModelClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            If _tabs.SelectedIndex = BomTabIndex AndAlso _bomGrid.SelectedRows.Count > 0 Then
                Dim selectedRow = _bomGrid.SelectedRows(0)
                Dim rowModel = Convert.ToString(selectedRow.Cells("FileName").Value)
                If String.IsNullOrWhiteSpace(rowModel) Then
                    rowModel = Convert.ToString(selectedRow.Cells("DisplayName").Value)
                End If
                If Not String.IsNullOrWhiteSpace(rowModel) Then
                    modelInput = rowModel
                End If
            End If

            If Not ValidateModelInput(modelInput) Then Return

            Dim formatName = Convert.ToString(_exportFormatComboBox.SelectedItem)
            Dim format = CType([Enum].Parse(GetType(ModelExportFormat), formatName, True), ModelExportFormat)
            Await RunOperationAsync(Function() _service.ExportModelAsync(modelInput, format), "Export " & formatName & " (" & modelInput & ")")
        End Sub

        Private Async Sub LoadParametersClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            If Not ValidateModelInput(modelInput) Then Return
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
                _paramsModelInput = modelInput

                _tabs.SelectedIndex = 1
                AddLog("OK: " & result.Message)
            Catch ex As Exception
                HandleUnexpectedError("Load Parameters", ex)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Async Sub LoadFeaturesClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            If Not ValidateModelInput(modelInput) Then Return
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
                HandleUnexpectedError("Load Features", ex)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Async Sub LoadComponentsClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            If Not ValidateModelInput(modelInput) Then Return
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
                HandleUnexpectedError("Load Components", ex)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Async Sub BuildBomClickedAsync(sender As Object, e As EventArgs)
            Dim modelInput = _modelNameTextBox.Text
            If Not ValidateModelInput(modelInput) Then Return
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
                _bomModelInput = modelInput
                _bomTopFileName = If(result.Rows.Count > 0, result.Rows(0).FileName, modelInput)

                AddLog("OK: " & result.Message)
                Await LoadRulesForCurrentBomAsync()

                RenderBomGrid()
                UpdateShowParametersButtonState()
                _itemParamsPanel.Visible = False
                WindowState = FormWindowState.Maximized

                _tabs.SelectedIndex = 2
            Catch ex As Exception
                HandleUnexpectedError("Build BOM", ex)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Async Function LoadRulesForCurrentBomAsync() As Task
            Dim pathResult = Await ResolveRulesPathAsync()
            If Not pathResult.Item1 Then
                AddLog("WARN: BOM rules not loaded. " & pathResult.Item2)
                Return
            End If

            If Not File.Exists(pathResult.Item2) Then
                AddLog("No saved BOM rules found (" & Path.GetFileName(pathResult.Item2) & ").")
                Return
            End If

            Try
                Dim ruleFile = BomRulesStore.Load(pathResult.Item2)
                Dim warnings As New List(Of String)()
                Dim applied = BomRulesStore.ApplyToRows(ruleFile, _bomRowsAll, warnings)
                For Each warning In warnings
                    AddLog("WARN: " & warning)
                Next
                AddLog("OK: Loaded " & applied.ToString() & " BOM rule(s) from " & Path.GetFileName(pathResult.Item2) & ". Click Apply Config to evaluate.")
            Catch ex As BomRuleException
                AddLog("ERROR: " & ex.Message)
            Catch ex As Exception
                HandleUnexpectedError("Load BOM rules", ex, "Unable to read the BOM rules file.")
            End Try
        End Function

        Private Async Function ResolveRulesPathAsync() As Task(Of Tuple(Of Boolean, String))
            Dim dirResult = Await _service.GetWorkingDirectoryAsync()
            If Not dirResult.Success Then
                Return Tuple.Create(False, dirResult.Message)
            End If

            Return BomRulesStore.BuildRulesPath(dirResult.Message, _bomTopFileName)
        End Function

        Private Async Sub ApplyConfigClicked(sender As Object, e As EventArgs)
            _bomGrid.EndEdit()
            If _bomRowsAll.Count = 0 Then
                AddLog("ERROR: Build the BOM before applying the configuration.")
                Return
            End If

            If _paramsGrid.Rows.Count = 0 Then
                AddLog("ERROR: Load Params for the top-level model before applying the configuration.")
                Return
            End If

            If Not String.Equals(NormalizeModelKey(_paramsModelInput), NormalizeModelKey(_bomModelInput), StringComparison.OrdinalIgnoreCase) Then
                AddLog("WARN: Parameters were loaded for '" & _paramsModelInput & "' but the BOM is for '" & _bomModelInput & "'.")
            End If

            Dim selectedParams As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each gridRow As DataGridViewRow In _paramsGrid.Rows
                If gridRow.IsNewRow Then
                    Continue For
                End If

                Dim paramName = Convert.ToString(gridRow.Cells("Name").Value)
                If String.IsNullOrWhiteSpace(paramName) Then
                    Continue For
                End If

                Dim newValue = Convert.ToString(gridRow.Cells("NewValue").Value)
                selectedParams(paramName.Trim()) = If(String.IsNullOrWhiteSpace(newValue), Convert.ToString(gridRow.Cells("Value").Value), newValue.Trim())
            Next

            Try
                Dim summary = BomConfigEvaluator.Evaluate(_bomRowsAll, selectedParams)
                _bomRulesDirty = False
                RenderBomGrid()

                Const MaxLoggedErrors As Integer = 20
                Dim logged = 0
                For Each row In _bomRowsAll
                    If logged >= MaxLoggedErrors Then
                        AddLog("WARN: ... more rule errors (see red cells).")
                        Exit For
                    End If
                    If row.ConditionError IsNot Nothing Then
                        AddLog("WARN: " & row.FileName & " [" & row.FeaturePath & "] Condition: " & row.ConditionError)
                        logged += 1
                    End If
                    If row.InheritanceError IsNot Nothing Then
                        AddLog("WARN: " & row.FileName & " [" & row.FeaturePath & "] Inheritance: " & row.InheritanceError)
                        logged += 1
                    End If
                Next

                AddLog("OK: Config applied - " & summary.IncludedCount.ToString() & " included, " &
                       summary.ExcludedCount.ToString() & " excluded, " & summary.ErrorCount.ToString() & " row(s) with rule errors.")

                Await SyncComponentVisibilityAsync()
            Catch ex As Exception
                HandleUnexpectedError("Apply Config", ex)
            End Try
        End Sub

        ' Every component is sent so each owning assembly is unhidden first; only the topmost excluded ones are re-hidden.
        Private Async Function SyncComponentVisibilityAsync() As Task
            Dim items As New List(Of ComponentVisibilityItem)()
            For Each row In _bomRowsAll
                If row.ParentNodeId = 0 OrElse String.IsNullOrEmpty(row.FeaturePath) Then
                    Continue For
                End If

                Dim parentRow As BomGridRowItem = Nothing
                Dim parentExcluded = _bomRowsByNodeId.TryGetValue(row.ParentNodeId, parentRow) AndAlso parentRow.IsExcluded

                items.Add(New ComponentVisibilityItem With {
                    .FeaturePath = row.FeaturePath,
                    .Hide = row.IsExcluded AndAlso Not parentExcluded,
                    .Label = row.FileName & " [" & row.FeaturePath & "]"
                })
            Next

            If items.Count = 0 Then
                Return
            End If

            SetBusyState(True)
            Try
                Dim result = Await _service.ApplyComponentVisibilityAsync(_bomModelInput, items)
                AddLog(If(result.Success, "OK: ", "WARN: ") & result.Message)
            Finally
                SetBusyState(False)
            End Try
        End Function

        Private Async Sub SaveRulesClickedAsync(sender As Object, e As EventArgs)
            _bomGrid.EndEdit()
            If _bomRowsAll.Count = 0 Then
                AddLog("ERROR: Build the BOM before saving rules.")
                Return
            End If

            SetBusyState(True)
            Try
                Dim pathResult = Await ResolveRulesPathAsync()
                If Not pathResult.Item1 Then
                    AddLog("ERROR: " & pathResult.Item2)
                    Return
                End If

                Dim ruleFile = BomRulesStore.CreateFromRows(_bomTopFileName, _bomRowsAll)
                BomRulesStore.Save(pathResult.Item2, ruleFile)
                AddLog("OK: Saved " & ruleFile.Rules.Count.ToString() & " BOM rule(s) to " & pathResult.Item2)
            Catch ex As Exception
                HandleUnexpectedError("Save BOM rules", ex, "Unable to save the BOM rules file.")
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Async Sub PushToModelsClickedAsync(sender As Object, e As EventArgs)
            _bomGrid.EndEdit()
            If _bomRowsAll.Count = 0 Then
                AddLog("ERROR: Build the BOM before pushing values to models.")
                Return
            End If

            If _bomRulesDirty Then
                AddLog("ERROR: Rules changed since the last Apply Config. Click Apply Config first.")
                Return
            End If

            ' One entry per file: identical values are pushed once, conflicting values are skipped.
            Dim firstPathByFile As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim valuesByFile As New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.OrdinalIgnoreCase)
            Dim conflictsByFile As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
            Dim fileOrder As New List(Of String)()

            For Each row In _bomRowsAll
                If row.ParentNodeId = 0 OrElse row.IsExcluded OrElse row.InheritanceError IsNot Nothing OrElse
                   row.ResolvedAssignments Is Nothing OrElse row.ResolvedAssignments.Count = 0 OrElse
                   String.IsNullOrEmpty(row.FeaturePath) OrElse String.IsNullOrEmpty(row.FileName) Then
                    Continue For
                End If

                Dim fileValues As Dictionary(Of String, String) = Nothing
                If Not valuesByFile.TryGetValue(row.FileName, fileValues) Then
                    fileValues = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                    valuesByFile(row.FileName) = fileValues
                    conflictsByFile(row.FileName) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                    firstPathByFile(row.FileName) = row.FeaturePath
                    fileOrder.Add(row.FileName)
                End If

                For Each assignment In row.ResolvedAssignments
                    Dim existing As String = Nothing
                    If Not fileValues.TryGetValue(assignment.Key, existing) Then
                        fileValues(assignment.Key) = assignment.Value
                    ElseIf Not String.Equals(existing, assignment.Value, StringComparison.Ordinal) Then
                        conflictsByFile(row.FileName).Add(assignment.Key)
                    End If
                Next
            Next

            Dim items As New List(Of InheritedParameterItem)()
            Dim parameterCount = 0
            For Each fileName In fileOrder
                Dim conflicts = conflictsByFile(fileName)
                For Each conflictName In conflicts
                    AddLog("WARN: " & fileName & ": instances resolve different values for " & conflictName & "; skipped.")
                Next

                Dim assignments As New List(Of KeyValuePair(Of String, String))()
                For Each pair In valuesByFile(fileName)
                    If Not conflicts.Contains(pair.Key) Then
                        assignments.Add(pair)
                    End If
                Next

                If assignments.Count = 0 Then
                    Continue For
                End If

                parameterCount += assignments.Count
                items.Add(New InheritedParameterItem With {
                    .FeaturePath = firstPathByFile(fileName),
                    .Label = fileName,
                    .Assignments = assignments
                })
            Next

            If items.Count = 0 Then
                AddLog("Nothing to push. Click Apply Config first.")
                Return
            End If

            Dim prompt = "Push " & parameterCount.ToString() & " inherited parameter value(s) to " & items.Count.ToString() & " model(s)?" &
                         Environment.NewLine & Environment.NewLine &
                         "Changes stay in the Creo session until you save the models in Creo."
            If MessageBox.Show(Me, prompt, "Push to Models", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then
                Return
            End If

            SetBusyState(True)
            Try
                Dim result = Await _service.PushInheritedParametersAsync(_bomModelInput, items)
                AddLog(If(result.Success, "OK: ", "WARN: ") & result.Message)
            Catch ex As Exception
                HandleUnexpectedError("Push to Models", ex)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Sub HideExcludedClicked(sender As Object, e As EventArgs)
            _bomGrid.EndEdit()
            _hideExcluded = Not _hideExcluded
            _hideExcludedButton.Text = If(_hideExcluded, "Show Excluded", "Hide Excluded")
            RenderBomGrid()
        End Sub

        Private Sub BomGridCellEndEdit(sender As Object, e As DataGridViewCellEventArgs)
            If e.RowIndex < 0 Then
                Return
            End If

            Dim columnName = _bomGrid.Columns(e.ColumnIndex).Name
            If columnName <> "Conditions" AndAlso columnName <> "Inheritance" Then
                Return
            End If

            Dim nodeIdObj = _bomGrid.Rows(e.RowIndex).Tag
            Dim row As BomGridRowItem = Nothing
            If nodeIdObj Is Nothing OrElse Not _bomRowsByNodeId.TryGetValue(CInt(nodeIdObj), row) Then
                Return
            End If

            Dim ruleText = Convert.ToString(_bomGrid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
            ruleText = If(ruleText, String.Empty).Replace(vbCr, " ").Replace(vbLf, " ").Replace(vbTab, " ").Trim()
            If columnName = "Conditions" Then
                If Not String.Equals(If(row.Conditions, String.Empty), ruleText, StringComparison.Ordinal) Then _bomRulesDirty = True
                row.Conditions = ruleText
            Else
                If Not String.Equals(If(row.Inheritance, String.Empty), ruleText, StringComparison.Ordinal) Then _bomRulesDirty = True
                row.Inheritance = ruleText
            End If
        End Sub

        Private Shared Function NormalizeModelKey(modelInput As String) As String
            If String.IsNullOrWhiteSpace(modelInput) Then
                Return String.Empty
            End If

            Try
                Return ModelNameParser.Parse(modelInput).Name
            Catch
                Return modelInput.Trim()
            End Try
        End Function

        Private Async Sub SetSelectedParameterClickedAsync(sender As Object, e As EventArgs)
            Dim updates = CollectPendingParameterUpdates(_paramsGrid)
            If updates.Count = 0 Then
                AddLog("ERROR: Enter at least one value in the 'New Value' column before setting.")
                Return
            End If

            Dim modelInput = _modelNameTextBox.Text
            If Not ValidateModelInput(modelInput) Then Return
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
                _paramsModelInput = modelInput
                _tabs.SelectedIndex = 1
                AddLog("OK: " & result.Message)
            Catch ex As Exception
                HandleUnexpectedError("Set Parameters", ex)
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
                HandleUnexpectedError(actionName, ex)
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

                    Dim rowIndex = _bomGrid.Rows.Add(toggleText, CreateIndentedLevel(row.Level), row.FeatureId.ToString(), row.DisplayName, row.FileName, row.Descp, row.BsClass, row.Conditions, row.Inheritance, row.ResolvedInheritance)
                    Dim gridRow = _bomGrid.Rows(rowIndex)
                    gridRow.Tag = row.NodeId
                    ApplyBomRowStyle(gridRow, row)
                Next
            Finally
                _suppressSelectionChanged = False
            End Try

            UpdateShowParametersButtonState()
        End Sub

        Private Sub ApplyBomRowStyle(gridRow As DataGridViewRow, row As BomGridRowItem)
            If row.IsExcluded Then
                gridRow.DefaultCellStyle.Font = _strikeoutFont
                gridRow.DefaultCellStyle.ForeColor = Color.Gray
            End If

            Dim conditionsCell = gridRow.Cells("Conditions")
            Dim inheritanceCell = gridRow.Cells("Inheritance")
            Dim resolvedCell = gridRow.Cells("Resolved")
            If row.ParentNodeId = 0 Then
                conditionsCell.ReadOnly = True
                inheritanceCell.ReadOnly = True
                conditionsCell.Style.BackColor = Color.Gainsboro
                inheritanceCell.Style.BackColor = Color.Gainsboro
                resolvedCell.Style.BackColor = Color.Gainsboro
                conditionsCell.ToolTipText = "Top-level model uses the Parameters tab values."
                Return
            End If

            If Not String.IsNullOrEmpty(row.ResolvedInheritance) Then
                resolvedCell.ToolTipText = row.ResolvedInheritance
            End If

            ApplyRuleErrorStyle(conditionsCell, row.ConditionError)
            ApplyRuleErrorStyle(inheritanceCell, row.InheritanceError)
        End Sub

        Private Shared Sub ApplyRuleErrorStyle(cell As DataGridViewCell, errorText As String)
            If String.IsNullOrEmpty(errorText) Then
                Return
            End If

            cell.Style.BackColor = Color.MistyRose
            cell.Style.ForeColor = Color.DarkRed
            cell.ToolTipText = errorText
        End Sub

        Private Async Sub BomGridSelectionChanged(sender As Object, e As EventArgs)
            UpdateShowParametersButtonState()
            If _suppressSelectionChanged Then Return
            If Not _itemParamsPanel.Visible Then Return
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
            If _tabs.SelectedIndex <> BomTabIndex Then
                _itemParamsPanel.Visible = False
            End If
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
            _itemParamsPanel.Visible = True
            UpdateItemPanelWidth()
            Await LoadItemParametersForSelectedRowAsync(forceReload:=True)
        End Sub

        Private Sub CloseItemParametersPanel(sender As Object, e As EventArgs)
            _itemParamsPanel.Visible = False
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
                HandleUnexpectedError("Load Item Parameters", ex)
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
                HandleUnexpectedError("Set Item Parameters", ex)
            Finally
                SetBusyState(False)
            End Try
        End Sub

        Private Function IsRowVisible(row As BomGridRowItem) As Boolean
            If _hideExcluded AndAlso row.IsExcluded Then
                Return False
            End If

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
                dialog.DefaultExt = "csv"
                dialog.AddExtension = True
                dialog.CheckPathExists = True
                dialog.OverwritePrompt = True

                If dialog.ShowDialog(Me) <> DialogResult.OK Then
                    Return
                End If

                Dim validationResult = ValidateExportPath(dialog.FileName)
                If Not validationResult.Item1 Then
                    AddLog("ERROR: " & validationResult.Item2)
                    Return
                End If

                Try
                    Using writer As New StreamWriter(validationResult.Item2, False, Encoding.UTF8)
                        Dim headers As String() = {"Level", "Feature ID", "Name", "File Name", "Descp", "BS_Class", "Conditions", "Inheritance", "Resolved Values", "Included"}
                        writer.WriteLine(String.Join(",", EscapeCsvValues(headers)))

                        For Each row As DataGridViewRow In _bomGrid.Rows
                            If row.IsNewRow Then
                                Continue For
                            End If

                            Dim bomRow As BomGridRowItem = Nothing
                            Dim isExcluded = row.Tag IsNot Nothing AndAlso _bomRowsByNodeId.TryGetValue(CInt(row.Tag), bomRow) AndAlso bomRow.IsExcluded

                            Dim values As String() = {
                                Convert.ToString(row.Cells("Level").Value),
                                Convert.ToString(row.Cells("FeatureId").Value),
                                Convert.ToString(row.Cells("DisplayName").Value),
                                Convert.ToString(row.Cells("FileName").Value),
                                Convert.ToString(row.Cells("Descp").Value),
                                Convert.ToString(row.Cells("BsClass").Value),
                                Convert.ToString(row.Cells("Conditions").Value),
                                Convert.ToString(row.Cells("Inheritance").Value),
                                Convert.ToString(row.Cells("Resolved").Value),
                                If(isExcluded, "No", "Yes")
                            }
                            writer.WriteLine(String.Join(",", EscapeCsvValues(values)))
                        Next
                    End Using

                    AddLog("OK: BOM CSV exported.")
                Catch ex As Exception
                    HandleUnexpectedError("BOM CSV export", ex, "Unable to export the BOM CSV.")
                End Try
            End Using
        End Sub

        Private Function EscapeCsvValues(values As IEnumerable(Of String)) As IEnumerable(Of String)
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

        Private Sub SetBusyState(isBusy As Boolean)
            UseWaitCursor = isBusy
            _connectButton.Enabled = Not isBusy
            _retrieveButton.Enabled = Not isBusy
            _regenerateButton.Enabled = Not isBusy
            _exportFormatComboBox.Enabled = Not isBusy
            _exportModelButton.Enabled = Not isBusy
            _loadParametersButton.Enabled = Not isBusy
            _loadFeaturesButton.Enabled = Not isBusy
            _loadComponentsButton.Enabled = Not isBusy
            _buildBomButton.Enabled = Not isBusy
            _exportBomButton.Enabled = Not isBusy
            _expandAllBomButton.Enabled = Not isBusy
            _collapseAllBomButton.Enabled = Not isBusy
            _applyConfigButton.Enabled = Not isBusy
            _saveRulesButton.Enabled = Not isBusy
            _pushToModelsButton.Enabled = Not isBusy
            _hideExcludedButton.Enabled = Not isBusy
            _showItemParametersButton.Enabled = Not isBusy AndAlso _tabs.SelectedIndex = BomTabIndex AndAlso _bomGrid.SelectedRows.Count > 0
            _setParameterButton.Enabled = Not isBusy
            _clearLogButton.Enabled = Not isBusy
            _activateButton.Enabled = Not isBusy
            _itemSetParameterButton.Enabled = Not isBusy
            _itemParamsGrid.Enabled = Not isBusy

            If Not isBusy Then
                ResetGridCursors()
            End If
        End Sub

        ' DataGridView caches Cursor (WaitCursor while busy) on column-border hover and restores it later.
        Private Sub ResetGridCursors()
            For Each grid In New DataGridView() {_paramsGrid, _featuresGrid, _componentsGrid, _bomGrid, _itemParamsGrid}
                grid.Cursor = Cursors.Default
                If grid.EditingControl IsNot Nothing Then
                    grid.EditingControl.UseWaitCursor = False
                End If
            Next
        End Sub

        Private Shared Function BuildWindowTitle(status As LicenseStatus) As String
            If status Is Nothing OrElse status.State = LicenseState.Licensed Then
                Return "BOMForge"
            End If

            Return "BOMForge - Trial: " & status.DaysRemaining.ToString() &
                   If(status.DaysRemaining = 1, " day", " days") & " remaining"
        End Function

        Private Sub ActivateClicked(sender As Object, e As EventArgs)
            Using dialog As New ActivationForm(_licenseStatus)
                dialog.ShowDialog(Me)

                If dialog.ResultStatus.State <> LicenseState.Licensed Then
                    Return
                End If

                _licenseStatus = dialog.ResultStatus
                Text = BuildWindowTitle(_licenseStatus)
                _activateButton.Visible = False
                AddLog("OK: BOMForge is licensed to " & _licenseStatus.Licensee & ".")
            End Using
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

            LicenseManager.RecordHeartbeat()

            Try
                _service.Dispose()
            Catch
            End Try
        End Sub
    End Class
End Namespace
