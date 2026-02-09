using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Editor form for managing pick lists with Tree mode (hierarchical) or Builder mode (sentence construction).
/// </summary>
public class PickListEditorForm : Form
{
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        try
        {
            int value = 1;
            DwmSetWindowAttribute(Handle, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
        }
        catch { }
    }

    private readonly Configuration _config;
    private List<PickListConfig> _pickLists;
    private PickListConfig? _selectedList;
    private PickListNode? _selectedNode;
    private PickListCategory? _selectedCategory;
    private int _selectedOptionIndex = -1;
    private bool _suppressEvents;

    // Left panel
    private ListBox _listBox = null!;
    private TextBox _listNameBox = null!;
    private ComboBox _modeCombo = null!;
    private TextBox _listCriteriaRequiredBox = null!;
    private TextBox _listCriteriaAnyOfBox = null!;
    private TextBox _listCriteriaExcludeBox = null!;
    private CheckBox _listEnabledCheck = null!;

    // Tree style controls
    private Panel _treeStylePanel = null!;
    private RadioButton _freeformStyleRadio = null!;
    private RadioButton _structuredStyleRadio = null!;
    private ToolTip _styleToolTip = null!;
    private Label _structuredHelpLabel = null!;
    private Panel _structuredOptionsPanel = null!;
    private ComboBox _textPlacementCombo = null!;
    private CheckBox _blankLinesCheck = null!;

    // Tree mode panel (right side)
    private Panel _treeModePanel = null!;
    private TreeView _treeView = null!;
    private Button _addNodeBtn = null!;
    private Button _addChildBtn = null!;
    private Button _cloneNodeBtn = null!;
    private Button _removeNodeBtn = null!;
    private Button _moveUpBtn = null!;
    private Button _moveDownBtn = null!;
    private TextBox _nodeLabelBox = null!;
    private TextBox _nodeTextBox = null!;
    private Label _breadcrumbLabel = null!;
    private Label _selectedNodeHeader = null!;
    private Label _labelLabel = null!;
    private Label _textToInsertLabel = null!;

    // Builder/Macro reference UI for tree nodes
    private RadioButton _nodeTypeTextRadio = null!;
    private RadioButton _nodeTypeBuilderRadio = null!;
    private RadioButton _nodeTypeMacroRadio = null!;
    private ComboBox _nodeRefCombo = null!;  // Single combo for both builder and macro refs
    private Label _nodeRefLabel = null!;     // Single label for both
    private Label _prefixTextLabel = null!;
    private Panel _nodeTypePanel = null!;

    // Builder mode panel (right side)
    private Panel _builderModePanel = null!;
    private ListBox _categoryListBox = null!;
    private ListBox _optionsListBox = null!;
    private Button _addCategoryBtn = null!;
    private Button _removeCategoryBtn = null!;
    private Button _moveCategoryUpBtn = null!;
    private Button _moveCategoryDownBtn = null!;
    private Button _addOptionBtn = null!;
    private Button _removeOptionBtn = null!;
    private Button _moveOptionUpBtn = null!;
    private Button _moveOptionDownBtn = null!;
    private TextBox _categoryNameBox = null!;
    private TextBox _separatorBox = null!;
    private TextBox _optionTextBox = null!;
    private CheckBox _terminalCheckBox = null!;
    private Label _optionsHeader = null!;
    private Button _importBtn = null!;

    // Builder option reference UI
    private RadioButton _optTypeTextRadio = null!;
    private RadioButton _optTypeMacroRadio = null!;
    private RadioButton _optTypeTreeRadio = null!;
    private Panel _optTypePanel = null!;
    private ComboBox _optMacroRefCombo = null!;
    private ComboBox _optTreeRefCombo = null!;
    private Label _optMacroRefLabel = null!;
    private Label _optTreeRefLabel = null!;

    public PickListEditorForm(Configuration config)
    {
        _config = config;
        _pickLists = config.PickLists.Select(ClonePickList).ToList();
        InitializeUI();
        RefreshListBox();
    }

    private PickListConfig ClonePickList(PickListConfig pl)
    {
        return new PickListConfig
        {
            Id = pl.Id,
            Enabled = pl.Enabled,
            Name = pl.Name,
            Mode = pl.Mode,
            TreeStyle = pl.TreeStyle,
            StructuredTextPlacement = pl.StructuredTextPlacement,
            StructuredBlankLines = pl.StructuredBlankLines,
            CriteriaRequired = pl.CriteriaRequired,
            CriteriaAnyOf = pl.CriteriaAnyOf,
            CriteriaExclude = pl.CriteriaExclude,
            Nodes = pl.Nodes.Select(n => n.Clone()).ToList(),
            Categories = pl.Categories.Select(c => new PickListCategory
            {
                Name = c.Name,
                Options = c.Options.Select(o => new BuilderOption
                {
                    Text = o.Text,
                    MacroId = o.MacroId,
                    TreeListId = o.TreeListId
                }).ToList(),
                OptionsLegacy = new List<string>(c.OptionsLegacy),
                Separator = c.Separator,
                TerminalOptions = new List<int>(c.TerminalOptions)
            }).ToList()
        };
    }

    private void InitializeUI()
    {
        Text = "Pick List Editor";
        Size = new Size(1000, 650);
        MinimumSize = new Size(900, 550);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(30, 30, 30);
        ForeColor = Color.White;
        Padding = new Padding(15);

        // Bottom buttons
        var saveBtn = new Button
        {
            Text = "Save && Close",
            Size = new Size(100, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            BackColor = Color.FromArgb(51, 102, 0),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        saveBtn.FlatAppearance.BorderSize = 0;
        saveBtn.Click += (s, e) => SaveAndClose();
        Controls.Add(saveBtn);

        var cancelBtn = new Button
        {
            Text = "Cancel",
            Size = new Size(80, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
            BackColor = Color.FromArgb(102, 51, 0),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        cancelBtn.FlatAppearance.BorderSize = 0;
        cancelBtn.Click += (s, e) => Close();
        Controls.Add(cancelBtn);

        // Example buttons (bottom left)
        var treeExampleBtn = new Button
        {
            Text = "Tree Example",
            Size = new Size(95, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
            BackColor = Color.FromArgb(50, 70, 90),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        treeExampleBtn.FlatAppearance.BorderSize = 0;
        treeExampleBtn.Click += (s, e) => AddExampleTreeList();
        Controls.Add(treeExampleBtn);

        var builderExampleBtn = new Button
        {
            Text = "Builder Example",
            Size = new Size(110, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
            BackColor = Color.FromArgb(70, 50, 90),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        builderExampleBtn.FlatAppearance.BorderSize = 0;
        builderExampleBtn.Click += (s, e) => AddExampleBuilderList();
        Controls.Add(builderExampleBtn);

        // Backup/Restore buttons
        var backupBtn = new Button
        {
            Text = "Backup...",
            Size = new Size(75, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        backupBtn.FlatAppearance.BorderSize = 0;
        backupBtn.Click += (s, e) => BackupPickLists();
        Controls.Add(backupBtn);

        var restoreBtn = new Button
        {
            Text = "Restore...",
            Size = new Size(75, 32),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left,
            BackColor = Color.FromArgb(60, 60, 60),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand
        };
        restoreBtn.FlatAppearance.BorderSize = 0;
        restoreBtn.Click += (s, e) => RestorePickLists();
        Controls.Add(restoreBtn);

        // Position buttons
        Resize += (s, e) =>
        {
            cancelBtn.Location = new Point(ClientSize.Width - cancelBtn.Width - 15, ClientSize.Height - cancelBtn.Height - 15);
            saveBtn.Location = new Point(cancelBtn.Left - saveBtn.Width - 10, cancelBtn.Top);
            treeExampleBtn.Location = new Point(15, ClientSize.Height - treeExampleBtn.Height - 15);
            builderExampleBtn.Location = new Point(treeExampleBtn.Right + 10, treeExampleBtn.Top);
            backupBtn.Location = new Point(builderExampleBtn.Right + 20, treeExampleBtn.Top);
            restoreBtn.Location = new Point(backupBtn.Right + 5, treeExampleBtn.Top);
        };
        cancelBtn.Location = new Point(ClientSize.Width - cancelBtn.Width - 15, ClientSize.Height - cancelBtn.Height - 15);
        saveBtn.Location = new Point(cancelBtn.Left - saveBtn.Width - 10, cancelBtn.Top);
        treeExampleBtn.Location = new Point(15, ClientSize.Height - treeExampleBtn.Height - 15);
        builderExampleBtn.Location = new Point(treeExampleBtn.Right + 10, treeExampleBtn.Top);
        backupBtn.Location = new Point(builderExampleBtn.Right + 20, treeExampleBtn.Top);
        restoreBtn.Location = new Point(backupBtn.Right + 5, treeExampleBtn.Top);

        // Create layout
        CreateLeftPanel();
        CreateTreeModePanel();
        CreateBuilderModePanel();

        // Handle resize
        Resize += (s, e) => UpdateLayout();
        UpdateLayout();
    }

    private void CreateLeftPanel()
    {
        int x = 15, y = 15;

        // Header
        var header = CreateLabel("PICK LISTS", x, y, true);
        Controls.Add(header);

        // Buttons
        var addBtn = CreateButton("+", x + 95, y - 3, 26);
        addBtn.Click += (s, e) => AddList();
        Controls.Add(addBtn);

        var removeBtn = CreateButton("-", x + 125, y - 3, 26);
        removeBtn.Click += (s, e) => RemoveList();
        Controls.Add(removeBtn);

        var cloneBtn = CreateButton("Clone", x + 155, y - 3, 50);
        cloneBtn.Click += (s, e) => CloneList();
        Controls.Add(cloneBtn);

        var moveListUpBtn = CreateButton("^", x + 210, y - 3, 26);
        moveListUpBtn.Click += (s, e) => MoveList(-1);
        Controls.Add(moveListUpBtn);

        var moveListDownBtn = CreateButton("v", x + 240, y - 3, 26);
        moveListDownBtn.Click += (s, e) => MoveList(1);
        Controls.Add(moveListDownBtn);
        y += 25;

        // List box
        _listBox = new ListBox
        {
            Location = new Point(x, y),
            Size = new Size(260, 140),
            BackColor = Color.FromArgb(45, 45, 45),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            DrawMode = DrawMode.OwnerDrawFixed,
            ItemHeight = 24
        };
        _listBox.DrawItem += ListBox_DrawItem;
        _listBox.SelectedIndexChanged += ListBox_SelectedIndexChanged;
        Controls.Add(_listBox);
        y += 150;

        // Properties header
        Controls.Add(CreateLabel("LIST PROPERTIES", x, y, true, 8));
        y += 22;

        // Enabled
        _listEnabledCheck = new CheckBox
        {
            Text = "Enabled",
            Location = new Point(x, y),
            AutoSize = true,
            ForeColor = Color.White
        };
        _listEnabledCheck.CheckedChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedList == null) return;
            _selectedList.Enabled = _listEnabledCheck.Checked;
            RefreshListBox();
        };
        Controls.Add(_listEnabledCheck);
        y += 26;

        // Name
        Controls.Add(CreateLabel("Name:", x, y + 2));
        _listNameBox = CreateTextBox(x + 60, y, 200);
        _listNameBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedList == null) return;
            _selectedList.Name = _listNameBox.Text;
            RefreshListBox();
        };
        Controls.Add(_listNameBox);
        y += 28;

        // Mode
        Controls.Add(CreateLabel("Mode:", x, y + 2));
        _modeCombo = new ComboBox
        {
            Location = new Point(x + 60, y),
            Width = 100,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _modeCombo.Items.AddRange(new object[] { "Tree", "Builder" });
        _modeCombo.SelectedIndex = 0;
        _modeCombo.SelectedIndexChanged += ModeCombo_SelectedIndexChanged;
        Controls.Add(_modeCombo);
        y += 28;

        // Tree Style panel (only visible for Tree mode)
        _treeStylePanel = new Panel
        {
            Location = new Point(x, y),
            Size = new Size(260, 50),
            BackColor = Color.Transparent,
            Visible = true
        };
        Controls.Add(_treeStylePanel);

        var styleLabel = new Label
        {
            Text = "Style:",
            Location = new Point(0, 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", 9)
        };
        _treeStylePanel.Controls.Add(styleLabel);

        _freeformStyleRadio = new RadioButton
        {
            Text = "Freeform",
            Location = new Point(60, 0),
            AutoSize = true,
            ForeColor = Color.FromArgb(100, 220, 150),
            Checked = true
        };
        _freeformStyleRadio.CheckedChanged += TreeStyleRadio_CheckedChanged;
        _treeStylePanel.Controls.Add(_freeformStyleRadio);

        // Help button for freeform
        var freeformHelp = new Label
        {
            Text = "?",
            Location = new Point(_freeformStyleRadio.Right - 2, 2),
            Size = new Size(14, 14),
            ForeColor = Color.FromArgb(100, 150, 200),
            Font = new Font("Segoe UI", 7, FontStyle.Bold),
            Cursor = Cursors.Help
        };
        _treeStylePanel.Controls.Add(freeformHelp);

        _structuredStyleRadio = new RadioButton
        {
            Text = "Structured",
            Location = new Point(60, 22),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 180, 100)
        };
        _structuredStyleRadio.CheckedChanged += TreeStyleRadio_CheckedChanged;
        _treeStylePanel.Controls.Add(_structuredStyleRadio);

        // Help button for structured
        _structuredHelpLabel = new Label
        {
            Text = "?",
            Location = new Point(_structuredStyleRadio.Right - 2, 24),
            Size = new Size(14, 14),
            ForeColor = Color.FromArgb(100, 150, 200),
            Font = new Font("Segoe UI", 7, FontStyle.Bold),
            Cursor = Cursors.Help
        };
        _treeStylePanel.Controls.Add(_structuredHelpLabel);

        // Tooltips for style help
        _styleToolTip = new ToolTip
        {
            AutoPopDelay = 15000,
            InitialDelay = 200,
            ReshowDelay = 100
        };
        _styleToolTip.SetToolTip(freeformHelp, "Freeform: Paste each selection immediately.\nResult: \"The lungs are clear. Heart size is normal.\"");
        _styleToolTip.SetToolTip(_freeformStyleRadio, "Each selection pastes immediately with leading space");
        _styleToolTip.SetToolTip(_structuredStyleRadio, "Accumulate selections, format as CATEGORY: text");
        UpdateStructuredTooltip();

        y += 52;

        // Structured formatting options panel (only visible when Structured is selected)
        _structuredOptionsPanel = new Panel
        {
            Location = new Point(x, y),
            Size = new Size(260, 52),
            BackColor = Color.Transparent,
            Visible = false
        };
        Controls.Add(_structuredOptionsPanel);

        var placementLabel = new Label
        {
            Text = "Text placement:",
            Location = new Point(20, 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", 8)
        };
        _structuredOptionsPanel.Controls.Add(placementLabel);

        _textPlacementCombo = new ComboBox
        {
            Location = new Point(110, 0),
            Width = 130,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _textPlacementCombo.Items.AddRange(new object[] { "Inline", "Below Heading" });
        _textPlacementCombo.SelectedIndex = 0;
        _textPlacementCombo.SelectedIndexChanged += TextPlacementCombo_SelectedIndexChanged;
        _structuredOptionsPanel.Controls.Add(_textPlacementCombo);

        _blankLinesCheck = new CheckBox
        {
            Text = "Blank line between sections",
            Location = new Point(20, 26),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 180, 180),
            Font = new Font("Segoe UI", 8)
        };
        _blankLinesCheck.CheckedChanged += BlankLinesCheck_CheckedChanged;
        _structuredOptionsPanel.Controls.Add(_blankLinesCheck);

        y += 55;

        // Required
        Controls.Add(CreateLabel("Required:", x, y + 2));
        _listCriteriaRequiredBox = CreateTextBox(x + 60, y, 200);
        _listCriteriaRequiredBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedList == null) return;
            _selectedList.CriteriaRequired = _listCriteriaRequiredBox.Text;
        };
        Controls.Add(_listCriteriaRequiredBox);
        y += 28;

        // Any of
        Controls.Add(CreateLabel("Any of:", x, y + 2));
        _listCriteriaAnyOfBox = CreateTextBox(x + 60, y, 200);
        _listCriteriaAnyOfBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedList == null) return;
            _selectedList.CriteriaAnyOf = _listCriteriaAnyOfBox.Text;
        };
        Controls.Add(_listCriteriaAnyOfBox);
        y += 28;

        // Exclude
        Controls.Add(CreateLabel("Exclude:", x, y + 2));
        _listCriteriaExcludeBox = CreateTextBox(x + 60, y, 200);
        _listCriteriaExcludeBox.TextChanged += (s, e) =>
        {
            if (_suppressEvents || _selectedList == null) return;
            _selectedList.CriteriaExclude = _listCriteriaExcludeBox.Text;
        };
        Controls.Add(_listCriteriaExcludeBox);
    }

    private void CreateTreeModePanel()
    {
        int x = 0, y = 0;

        _treeModePanel = new Panel
        {
            Location = new Point(295, 15),
            Size = new Size(680, 550),
            BackColor = Color.FromArgb(30, 30, 30)
        };
        Controls.Add(_treeModePanel);

        // Header
        var header = CreateLabel("NODES", x, y, true);
        _treeModePanel.Controls.Add(header);

        // Buttons
        _addNodeBtn = CreateButton("+ Add", x + 70, y - 3, 50);
        _addNodeBtn.Click += (s, e) => AddNode(false);
        _treeModePanel.Controls.Add(_addNodeBtn);

        _addChildBtn = CreateButton("+ Child", x + 125, y - 3, 55);
        _addChildBtn.Click += (s, e) => AddNode(true);
        _treeModePanel.Controls.Add(_addChildBtn);

        _cloneNodeBtn = CreateButton("Clone", x + 185, y - 3, 45);
        _cloneNodeBtn.Click += (s, e) => CloneNode();
        _treeModePanel.Controls.Add(_cloneNodeBtn);

        _moveUpBtn = CreateButton("^", x + 235, y - 3, 26);
        _moveUpBtn.Click += (s, e) => MoveNode(-1);
        _treeModePanel.Controls.Add(_moveUpBtn);

        _moveDownBtn = CreateButton("v", x + 265, y - 3, 26);
        _moveDownBtn.Click += (s, e) => MoveNode(1);
        _treeModePanel.Controls.Add(_moveDownBtn);

        _removeNodeBtn = CreateButton("- Del", x + 296, y - 3, 45);
        _removeNodeBtn.Click += (s, e) => RemoveNode();
        _treeModePanel.Controls.Add(_removeNodeBtn);
        y += 25;

        // Tree view
        _treeView = new TreeView
        {
            Location = new Point(x, y),
            Size = new Size(400, 280),
            BackColor = Color.FromArgb(45, 45, 45),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            FullRowSelect = true,
            HideSelection = false,
            ShowLines = true,
            ShowPlusMinus = true,
            Indent = 20,
            ItemHeight = 22
        };
        _treeView.AfterSelect += TreeView_AfterSelect;
        _treeModePanel.Controls.Add(_treeView);

        // Node properties header (position updated in UpdateLayout)
        _selectedNodeHeader = CreateLabel("SELECTED NODE", x, 0, true, 8);
        _treeModePanel.Controls.Add(_selectedNodeHeader);

        // Breadcrumb (position updated in UpdateLayout)
        _breadcrumbLabel = new Label
        {
            Location = new Point(x, 0),
            AutoSize = true,
            ForeColor = Color.FromArgb(100, 160, 200),
            Font = new Font("Segoe UI", 8)
        };
        _treeModePanel.Controls.Add(_breadcrumbLabel);

        // Label (position updated in UpdateLayout)
        _labelLabel = CreateLabel("Label:", x, 0);
        _treeModePanel.Controls.Add(_labelLabel);
        _nodeLabelBox = CreateTextBox(x + 50, 0, 300);
        _nodeLabelBox.TextChanged += NodeLabelBox_TextChanged;
        _treeModePanel.Controls.Add(_nodeLabelBox);

        // Node type panel (radio buttons for leaf type selection)
        _nodeTypePanel = new Panel
        {
            Location = new Point(x, 0),
            Size = new Size(400, 26),
            BackColor = Color.Transparent
        };
        _treeModePanel.Controls.Add(_nodeTypePanel);

        _nodeTypeTextRadio = new RadioButton
        {
            Text = "Text leaf",
            Location = new Point(0, 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(100, 220, 150),
            Checked = true
        };
        _nodeTypeTextRadio.CheckedChanged += NodeTypeRadio_CheckedChanged;
        _nodeTypePanel.Controls.Add(_nodeTypeTextRadio);

        _nodeTypeBuilderRadio = new RadioButton
        {
            Text = "Builder ref",
            Location = new Point(90, 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 180, 100)
        };
        _nodeTypeBuilderRadio.CheckedChanged += NodeTypeRadio_CheckedChanged;
        _nodeTypePanel.Controls.Add(_nodeTypeBuilderRadio);

        _nodeTypeMacroRadio = new RadioButton
        {
            Text = "Macro ref",
            Location = new Point(195, 2),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 150, 255)
        };
        _nodeTypeMacroRadio.CheckedChanged += NodeTypeRadio_CheckedChanged;
        _nodeTypePanel.Controls.Add(_nodeTypeMacroRadio);

        // Reference dropdown — shared by builder ref and macro ref (position updated in UpdateLayout)
        _nodeRefLabel = CreateLabel("Reference:", x, 0);
        _treeModePanel.Controls.Add(_nodeRefLabel);

        _nodeRefCombo = new ComboBox
        {
            Location = new Point(x + 80, 0),
            Width = 250,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            DropDownStyle = ComboBoxStyle.DropDownList,
            MaxDropDownItems = 12
        };
        _nodeRefCombo.SelectedIndexChanged += NodeRefCombo_SelectedIndexChanged;
        _treeModePanel.Controls.Add(_nodeRefCombo);

        // Prefix text label (for builder ref nodes)
        _prefixTextLabel = CreateLabel("Prefix text (optional):", x, 0);
        _prefixTextLabel.ForeColor = Color.FromArgb(255, 180, 100);
        _treeModePanel.Controls.Add(_prefixTextLabel);

        // Text to insert (position updated in UpdateLayout)
        _textToInsertLabel = CreateLabel("Text to insert:", x, 0);
        _treeModePanel.Controls.Add(_textToInsertLabel);

        _nodeTextBox = new TextBox
        {
            Location = new Point(x, 0),
            Size = new Size(400, 100),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };
        _nodeTextBox.TextChanged += NodeTextBox_TextChanged;
        _treeModePanel.Controls.Add(_nodeTextBox);
    }

    private void CreateBuilderModePanel()
    {
        int x = 0, y = 0;

        _builderModePanel = new Panel
        {
            Location = new Point(295, 15),
            Size = new Size(680, 550),
            BackColor = Color.FromArgb(30, 30, 30),
            Visible = false
        };
        Controls.Add(_builderModePanel);

        // Categories section
        var catHeader = CreateLabel("CATEGORIES", x, y, true);
        _builderModePanel.Controls.Add(catHeader);

        _addCategoryBtn = CreateButton("+", x + 100, y - 3, 26);
        _addCategoryBtn.Click += (s, e) => AddCategory();
        _builderModePanel.Controls.Add(_addCategoryBtn);

        _removeCategoryBtn = CreateButton("-", x + 130, y - 3, 26);
        _removeCategoryBtn.Click += (s, e) => RemoveCategory();
        _builderModePanel.Controls.Add(_removeCategoryBtn);

        _moveCategoryUpBtn = CreateButton("^", x + 160, y - 3, 26);
        _moveCategoryUpBtn.Click += (s, e) => MoveCategory(-1);
        _builderModePanel.Controls.Add(_moveCategoryUpBtn);

        _moveCategoryDownBtn = CreateButton("v", x + 190, y - 3, 26);
        _moveCategoryDownBtn.Click += (s, e) => MoveCategory(1);
        _builderModePanel.Controls.Add(_moveCategoryDownBtn);

        _importBtn = CreateButton("Import...", x + 225, y - 3, 60);
        _importBtn.Click += (s, e) => ImportCategories();
        _builderModePanel.Controls.Add(_importBtn);
        y += 25;

        _categoryListBox = new ListBox
        {
            Location = new Point(x, y),
            Size = new Size(280, 200),
            BackColor = Color.FromArgb(45, 45, 45),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            DrawMode = DrawMode.OwnerDrawFixed,
            ItemHeight = 24
        };
        _categoryListBox.DrawItem += CategoryListBox_DrawItem;
        _categoryListBox.SelectedIndexChanged += CategoryListBox_SelectedIndexChanged;
        _builderModePanel.Controls.Add(_categoryListBox);

        // Category properties
        var catPropY = y + 210;
        _builderModePanel.Controls.Add(CreateLabel("Name:", x, catPropY + 2));
        _categoryNameBox = CreateTextBox(x + 60, catPropY, 220);
        _categoryNameBox.TextChanged += CategoryNameBox_TextChanged;
        _builderModePanel.Controls.Add(_categoryNameBox);
        catPropY += 28;

        _builderModePanel.Controls.Add(CreateLabel("Separator:", x, catPropY + 2));
        _separatorBox = CreateTextBox(x + 60, catPropY, 60);
        _separatorBox.TextChanged += SeparatorBox_TextChanged;
        _builderModePanel.Controls.Add(_separatorBox);

        var sepHint = CreateLabel("(text after selection, default: space)", x + 130, catPropY + 2);
        sepHint.ForeColor = Color.FromArgb(100, 100, 100);
        _builderModePanel.Controls.Add(sepHint);

        // Options section (right side of builder panel)
        var optX = 310;
        _optionsHeader = CreateLabel("OPTIONS IN: (select category)", optX, 0, true);
        _builderModePanel.Controls.Add(_optionsHeader);

        _addOptionBtn = CreateButton("+", optX + 200, -3, 26);
        _addOptionBtn.Click += (s, e) => AddOption();
        _builderModePanel.Controls.Add(_addOptionBtn);

        _removeOptionBtn = CreateButton("-", optX + 230, -3, 26);
        _removeOptionBtn.Click += (s, e) => RemoveOption();
        _builderModePanel.Controls.Add(_removeOptionBtn);

        _moveOptionUpBtn = CreateButton("^", optX + 260, -3, 26);
        _moveOptionUpBtn.Click += (s, e) => MoveOption(-1);
        _builderModePanel.Controls.Add(_moveOptionUpBtn);

        _moveOptionDownBtn = CreateButton("v", optX + 290, -3, 26);
        _moveOptionDownBtn.Click += (s, e) => MoveOption(1);
        _builderModePanel.Controls.Add(_moveOptionDownBtn);

        _optionsListBox = new ListBox
        {
            Location = new Point(optX, y),
            Size = new Size(340, 200),
            BackColor = Color.FromArgb(45, 45, 45),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            DrawMode = DrawMode.OwnerDrawFixed,
            ItemHeight = 24
        };
        _optionsListBox.DrawItem += OptionsListBox_DrawItem;
        _optionsListBox.SelectedIndexChanged += OptionsListBox_SelectedIndexChanged;
        _builderModePanel.Controls.Add(_optionsListBox);

        // Option type panel
        var optPropY = y + 210;
        _optTypePanel = new Panel
        {
            Location = new Point(optX, optPropY),
            Size = new Size(340, 22),
            BackColor = Color.Transparent
        };
        _builderModePanel.Controls.Add(_optTypePanel);

        _optTypeTextRadio = new RadioButton
        {
            Text = "Text",
            Location = new Point(0, 0),
            AutoSize = true,
            ForeColor = Color.FromArgb(100, 220, 150),
            Checked = true,
            Font = new Font("Segoe UI", 8)
        };
        _optTypeTextRadio.CheckedChanged += OptTypeRadio_CheckedChanged;
        _optTypePanel.Controls.Add(_optTypeTextRadio);

        _optTypeMacroRadio = new RadioButton
        {
            Text = "Macro",
            Location = new Point(65, 0),
            AutoSize = true,
            ForeColor = Color.FromArgb(180, 150, 255),
            Font = new Font("Segoe UI", 8)
        };
        _optTypeMacroRadio.CheckedChanged += OptTypeRadio_CheckedChanged;
        _optTypePanel.Controls.Add(_optTypeMacroRadio);

        _optTypeTreeRadio = new RadioButton
        {
            Text = "Tree",
            Location = new Point(135, 0),
            AutoSize = true,
            ForeColor = Color.FromArgb(100, 180, 255),
            Font = new Font("Segoe UI", 8)
        };
        _optTypeTreeRadio.CheckedChanged += OptTypeRadio_CheckedChanged;
        _optTypePanel.Controls.Add(_optTypeTreeRadio);

        optPropY += 24;

        // Option text box (multiline for longer text)
        _builderModePanel.Controls.Add(CreateLabel("Option text:", optX, optPropY + 2));
        _optionTextBox = new TextBox
        {
            Location = new Point(optX, optPropY + 20),
            Size = new Size(340, 60),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical
        };
        _optionTextBox.TextChanged += OptionTextBox_TextChanged;
        _builderModePanel.Controls.Add(_optionTextBox);

        // Macro ref combo for options
        _optMacroRefLabel = CreateLabel("Macro:", optX, optPropY + 2);
        _optMacroRefLabel.ForeColor = Color.FromArgb(180, 150, 255);
        _optMacroRefLabel.Visible = false;
        _builderModePanel.Controls.Add(_optMacroRefLabel);

        _optMacroRefCombo = new ComboBox
        {
            Location = new Point(optX + 50, optPropY),
            Width = 290,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            DropDownStyle = ComboBoxStyle.DropDownList,
            MaxDropDownItems = 12,
            Visible = false
        };
        _optMacroRefCombo.SelectedIndexChanged += OptMacroRefCombo_SelectedIndexChanged;
        _builderModePanel.Controls.Add(_optMacroRefCombo);

        // Tree ref combo for options
        _optTreeRefLabel = CreateLabel("Tree:", optX, optPropY + 2);
        _optTreeRefLabel.ForeColor = Color.FromArgb(100, 180, 255);
        _optTreeRefLabel.Visible = false;
        _builderModePanel.Controls.Add(_optTreeRefLabel);

        _optTreeRefCombo = new ComboBox
        {
            Location = new Point(optX + 50, optPropY),
            Width = 290,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            DropDownStyle = ComboBoxStyle.DropDownList,
            MaxDropDownItems = 12,
            Visible = false
        };
        _optTreeRefCombo.SelectedIndexChanged += OptTreeRefCombo_SelectedIndexChanged;
        _builderModePanel.Controls.Add(_optTreeRefCombo);

        optPropY += 85;

        // Terminal checkbox
        _terminalCheckBox = new CheckBox
        {
            Text = "Completes sentence (skip remaining categories)",
            Location = new Point(optX, optPropY),
            AutoSize = true,
            ForeColor = Color.FromArgb(255, 200, 150)
        };
        _terminalCheckBox.CheckedChanged += TerminalCheckBox_CheckedChanged;
        _builderModePanel.Controls.Add(_terminalCheckBox);

        // Option hint
        var limitHint = CreateLabel("(keys 1-9 select first 9; click/arrow for rest)", optX, optPropY + 23);
        limitHint.ForeColor = Color.FromArgb(100, 100, 100);
        _builderModePanel.Controls.Add(limitHint);
    }

    private void UpdateLayout()
    {
        var w = ClientSize.Width;
        var h = ClientSize.Height;
        var rightX = 295;
        var rightW = w - rightX - 20;
        var bottomMargin = 60;

        // Update panel sizes
        _treeModePanel.Size = new Size(rightW, h - 30 - bottomMargin);
        _builderModePanel.Size = new Size(rightW, h - 30 - bottomMargin);

        if (_selectedList?.Mode == PickListMode.Tree)
        {
            UpdateTreeModeLayout(rightW, h - 30 - bottomMargin);
        }
        else
        {
            UpdateBuilderModeLayout(rightW, h - 30 - bottomMargin);
        }
    }

    private void UpdateTreeModeLayout(int panelW, int panelH)
    {
        // Tree view - takes upper portion
        var treeH = Math.Max(150, (panelH - 280) / 2 + 100);
        _treeView.Size = new Size(Math.Max(300, panelW - 10), treeH);

        // Position node properties section below tree view
        var y = _treeView.Bottom + 10;

        _selectedNodeHeader.Location = new Point(0, y);
        y += 20;

        _breadcrumbLabel.Location = new Point(0, y);
        y += 18;

        _labelLabel.Location = new Point(0, y + 2);
        _nodeLabelBox.Location = new Point(50, y);
        _nodeLabelBox.Width = Math.Max(200, panelW - 60);
        y += 28;

        // Node type selection (only visible for leaf nodes)
        _nodeTypePanel.Location = new Point(0, y);
        _nodeTypePanel.Width = Math.Max(300, panelW - 10);
        y += 28;

        // Reference dropdown (builder ref or macro ref)
        _nodeRefLabel.Location = new Point(0, y + 2);
        _nodeRefCombo.Location = new Point(80, y);
        _nodeRefCombo.Width = Math.Max(150, panelW - 100);
        y += 28;

        // Prefix text label (for builder ref) / Text to insert label (for text leaf)
        _prefixTextLabel.Location = new Point(0, y);
        _textToInsertLabel.Location = new Point(0, y);
        y += 20;

        var textBoxH = Math.Max(60, panelH - y - 10);
        _nodeTextBox.Location = new Point(0, y);
        _nodeTextBox.Size = new Size(Math.Max(300, panelW - 10), textBoxH);
    }

    private void UpdateBuilderModeLayout(int panelW, int panelH)
    {
        // Reserve space for option editing area at bottom (label + textbox + checkbox + hint = ~130px)
        var editAreaHeight = 130;
        var listH = Math.Max(120, panelH - editAreaHeight - 30);

        // Categories list - left half
        var catW = Math.Max(200, (panelW - 30) / 2);
        _categoryListBox.Size = new Size(catW, listH);

        // Options list - right half
        var optX = catW + 30;
        var optW = panelW - optX - 10;
        _optionsHeader.Location = new Point(optX, 0);
        _addOptionBtn.Location = new Point(optX + optW - 120, -3);
        _removeOptionBtn.Location = new Point(optX + optW - 90, -3);
        _moveOptionUpBtn.Location = new Point(optX + optW - 60, -3);
        _moveOptionDownBtn.Location = new Point(optX + optW - 30, -3);
        _optionsListBox.Location = new Point(optX, 25);
        _optionsListBox.Size = new Size(optW, listH);

        // Option editing area below options list
        var optPropY = _optionsListBox.Bottom + 10;

        // Option type panel
        _optTypePanel.Location = new Point(optX, optPropY);
        _optTypePanel.Width = optW;
        optPropY += 24;

        // "Option text:" label
        foreach (Control c in _builderModePanel.Controls)
        {
            if (c is Label lbl && lbl.Text == "Option text:")
                lbl.Location = new Point(optX, optPropY);
        }

        // Option text box
        _optionTextBox.Location = new Point(optX, optPropY + 18);
        _optionTextBox.Size = new Size(optW, 50);

        // Macro/Tree ref combos (same position as text box, shown/hidden based on type)
        _optMacroRefLabel.Location = new Point(optX, optPropY);
        _optMacroRefCombo.Location = new Point(optX + 50, optPropY);
        _optMacroRefCombo.Width = optW - 50;

        _optTreeRefLabel.Location = new Point(optX, optPropY);
        _optTreeRefCombo.Location = new Point(optX + 50, optPropY);
        _optTreeRefCombo.Width = optW - 50;

        // Terminal checkbox
        _terminalCheckBox.Location = new Point(optX, optPropY + 72);

        // Hint label
        foreach (Control c in _builderModePanel.Controls)
        {
            if (c is Label lbl && lbl.Text.StartsWith("(keys 1-9"))
                lbl.Location = new Point(optX, optPropY + 95);
        }
    }

    private void ModeCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedList == null) return;

        _selectedList.Mode = _modeCombo.SelectedIndex == 0 ? PickListMode.Tree : PickListMode.Builder;
        ShowModePanel();
        RefreshListBox();
    }

    private void TreeStyleRadio_CheckedChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedList == null) return;

        _selectedList.TreeStyle = _structuredStyleRadio.Checked
            ? TreePickListStyle.Structured
            : TreePickListStyle.Freeform;

        // Show/hide structured formatting options
        _structuredOptionsPanel.Visible = _structuredStyleRadio.Checked;
    }

    private void TextPlacementCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedList == null) return;

        _selectedList.StructuredTextPlacement = _textPlacementCombo.SelectedIndex == 1
            ? StructuredTextPlacement.BelowHeading
            : StructuredTextPlacement.Inline;

        UpdateStructuredTooltip();
    }

    private void BlankLinesCheck_CheckedChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedList == null) return;

        _selectedList.StructuredBlankLines = _blankLinesCheck.Checked;

        UpdateStructuredTooltip();
    }

    private void UpdateStructuredTooltip()
    {
        // Build example based on current settings
        var belowHeading = _textPlacementCombo?.SelectedIndex == 1;
        var blankLines = _blankLinesCheck?.Checked == true;

        string line1, line2;
        if (belowHeading)
        {
            line1 = "LUNGS:\n  Clear";
            line2 = "HEART:\n  Normal Size";
        }
        else
        {
            line1 = "LUNGS: Clear";
            line2 = "HEART: Normal Size";
        }

        var separator = blankLines ? "\n\n" : "\n";
        var example = line1 + separator + line2;

        var tooltip = $"Structured: Accumulate selections per category.\nResult:\n{example}\n\nUse Insert All to paste when done.";
        _styleToolTip?.SetToolTip(_structuredHelpLabel, tooltip);
    }

    private void ShowModePanel()
    {
        var isTreeMode = _selectedList?.Mode != PickListMode.Builder;
        _treeModePanel.Visible = isTreeMode;
        _builderModePanel.Visible = !isTreeMode;
        _treeStylePanel.Visible = isTreeMode;
        _structuredOptionsPanel.Visible = isTreeMode && _selectedList?.TreeStyle == TreePickListStyle.Structured;

        if (isTreeMode)
        {
            RefreshTreeView();
        }
        else
        {
            RefreshCategoryList();
        }

        UpdateLayout();
    }

    private Label CreateLabel(string text, int x, int y, bool bold = false, int size = 9)
    {
        return new Label
        {
            Text = text,
            Location = new Point(x, y),
            AutoSize = true,
            ForeColor = bold ? Color.FromArgb(180, 180, 180) : Color.FromArgb(150, 150, 150),
            Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular)
        };
    }

    private TextBox CreateTextBox(int x, int y, int width)
    {
        return new TextBox
        {
            Location = new Point(x, y),
            Width = width,
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
    }

    private Button CreateButton(string text, int x, int y, int width)
    {
        var btn = new Button
        {
            Text = text,
            Location = new Point(x, y),
            Size = new Size(width, 22),
            BackColor = Color.FromArgb(55, 55, 55),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 8)
        };
        btn.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
        return btn;
    }

    #region List Box (Pick Lists)

    private void ListBox_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= _pickLists.Count) return;

        var list = _pickLists[e.Index];
        e.DrawBackground();

        var isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        using var bgBrush = new SolidBrush(isSelected ? Color.FromArgb(50, 80, 110) : Color.FromArgb(45, 45, 45));
        e.Graphics.FillRectangle(bgBrush, e.Bounds);

        // Checkbox
        var checkRect = new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 5, 14, 14);
        using var checkBrush = new SolidBrush(list.Enabled ? Color.FromArgb(50, 120, 50) : Color.FromArgb(70, 70, 70));
        e.Graphics.FillRectangle(checkBrush, checkRect);
        if (list.Enabled)
        {
            using var pen = new Pen(Color.White, 2);
            e.Graphics.DrawLine(pen, checkRect.X + 3, checkRect.Y + 7, checkRect.X + 6, checkRect.Y + 10);
            e.Graphics.DrawLine(pen, checkRect.X + 6, checkRect.Y + 10, checkRect.X + 11, checkRect.Y + 4);
        }

        // Mode indicator
        var modeText = list.Mode == PickListMode.Builder ? "[B] " : "[T] ";
        var modeColor = list.Mode == PickListMode.Builder ? Color.FromArgb(255, 180, 100) : Color.FromArgb(100, 180, 255);

        // Name
        using var textBrush = new SolidBrush(list.Enabled ? modeColor : Color.Gray);
        var name = string.IsNullOrEmpty(list.Name) ? "(unnamed)" : list.Name;
        e.Graphics.DrawString(modeText + name, e.Font!, textBrush, e.Bounds.X + 26, e.Bounds.Y + 3);

        // Count
        var count = list.Mode == PickListMode.Builder
            ? list.Categories.Count
            : list.Nodes.Count + list.Nodes.Sum(n => n.CountDescendants());
        using var countBrush = new SolidBrush(Color.FromArgb(110, 110, 110));
        var countText = $"({count})";
        var countSize = e.Graphics.MeasureString(countText, e.Font!);
        e.Graphics.DrawString(countText, e.Font!, countBrush, e.Bounds.Right - countSize.Width - 8, e.Bounds.Y + 3);
    }

    private void ListBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_listBox.SelectedIndex >= 0 && _listBox.SelectedIndex < _pickLists.Count)
        {
            _selectedList = _pickLists[_listBox.SelectedIndex];
            LoadListProperties();
            ShowModePanel();
        }
        else
        {
            _selectedList = null;
            ClearListProperties();
            _treeView.Nodes.Clear();
            _categoryListBox.Items.Clear();
            _optionsListBox.Items.Clear();
        }
        UpdateButtonStates();
    }

    private void RefreshListBox()
    {
        var idx = _listBox.SelectedIndex;
        _listBox.Items.Clear();
        foreach (var list in _pickLists)
            _listBox.Items.Add(list);
        if (idx >= 0 && idx < _listBox.Items.Count)
            _listBox.SelectedIndex = idx;
        else if (_listBox.Items.Count > 0)
            _listBox.SelectedIndex = 0;
    }

    private void LoadListProperties()
    {
        if (_selectedList == null) return;
        _suppressEvents = true;
        try
        {
            _listEnabledCheck.Checked = _selectedList.Enabled;
            _listNameBox.Text = _selectedList.Name;
            _modeCombo.SelectedIndex = _selectedList.Mode == PickListMode.Builder ? 1 : 0;
            _freeformStyleRadio.Checked = _selectedList.TreeStyle == TreePickListStyle.Freeform;
            _structuredStyleRadio.Checked = _selectedList.TreeStyle == TreePickListStyle.Structured;
            _textPlacementCombo.SelectedIndex = _selectedList.StructuredTextPlacement == StructuredTextPlacement.BelowHeading ? 1 : 0;
            _blankLinesCheck.Checked = _selectedList.StructuredBlankLines;
            _structuredOptionsPanel.Visible = _selectedList.Mode == PickListMode.Tree && _selectedList.TreeStyle == TreePickListStyle.Structured;
            _listCriteriaRequiredBox.Text = _selectedList.CriteriaRequired;
            _listCriteriaAnyOfBox.Text = _selectedList.CriteriaAnyOf;
            _listCriteriaExcludeBox.Text = _selectedList.CriteriaExclude;
        }
        finally { _suppressEvents = false; }
        UpdateStructuredTooltip();
    }

    private void ClearListProperties()
    {
        _suppressEvents = true;
        try
        {
            _listEnabledCheck.Checked = false;
            _listNameBox.Text = "";
            _modeCombo.SelectedIndex = 0;
            _freeformStyleRadio.Checked = true;
            _textPlacementCombo.SelectedIndex = 0;
            _blankLinesCheck.Checked = false;
            _structuredOptionsPanel.Visible = false;
            _listCriteriaRequiredBox.Text = "";
            _listCriteriaAnyOfBox.Text = "";
            _listCriteriaExcludeBox.Text = "";
        }
        finally { _suppressEvents = false; }
    }

    #endregion

    #region Tree View (Tree Mode)

    private void RefreshTreeView()
    {
        _treeView.Nodes.Clear();
        _selectedNode = null;
        ClearNodeProperties();

        if (_selectedList == null) return;

        for (int i = 0; i < _selectedList.Nodes.Count; i++)
        {
            var node = _selectedList.Nodes[i];
            _treeView.Nodes.Add(CreateTreeNode(node, i + 1));
        }

        _treeView.ExpandAll();
        if (_treeView.Nodes.Count > 0)
            _treeView.SelectedNode = _treeView.Nodes[0];

        UpdateButtonStates();
    }

    private TreeNode CreateTreeNode(PickListNode node, int num)
    {
        var label = string.IsNullOrEmpty(node.Label) ? "(unnamed)" : node.Label;
        string suffix;
        Color color;

        if (node.HasChildren)
        {
            suffix = $" [{node.Children.Count}]";
            color = Color.FromArgb(100, 180, 255);  // Blue for branch
        }
        else if (node.IsMacroRef)
        {
            suffix = " [M]";
            color = Color.FromArgb(180, 150, 255);  // Purple for macro ref
        }
        else if (node.IsBuilderRef)
        {
            suffix = " [B]";
            color = Color.FromArgb(255, 180, 100);  // Orange for builder ref
        }
        else
        {
            suffix = !string.IsNullOrEmpty(node.Text) ? " *" : "";
            color = Color.White;  // White for text leaf
        }

        var treeNode = new TreeNode
        {
            Text = $"{num}. {label}{suffix}",
            Tag = node,
            ForeColor = color
        };

        for (int i = 0; i < node.Children.Count; i++)
            treeNode.Nodes.Add(CreateTreeNode(node.Children[i], i + 1));

        return treeNode;
    }

    private void TreeView_AfterSelect(object? sender, TreeViewEventArgs e)
    {
        _selectedNode = e.Node?.Tag as PickListNode;
        if (_selectedNode != null)
            LoadNodeProperties();
        else
            ClearNodeProperties();
        UpdateButtonStates();
    }

    private void LoadNodeProperties()
    {
        if (_selectedNode == null) return;
        _suppressEvents = true;
        try
        {
            _nodeLabelBox.Text = _selectedNode.Label;
            _nodeTextBox.Text = _selectedNode.Text;
            _breadcrumbLabel.Text = GetBreadcrumb();

            // Set node type radio buttons
            if (_selectedNode.IsMacroRef)
            {
                _nodeTypeMacroRadio.Checked = true;
                PopulateMacroRefCombo();
            }
            else if (_selectedNode.IsBuilderRef)
            {
                _nodeTypeBuilderRadio.Checked = true;
                PopulateBuilderRefCombo();
            }
            else
            {
                _nodeTypeTextRadio.Checked = true;
            }
        }
        finally { _suppressEvents = false; }
        UpdateNodeTypeUI();
    }

    private void ClearNodeProperties()
    {
        _suppressEvents = true;
        try
        {
            _nodeLabelBox.Text = "";
            _nodeTextBox.Text = "";
            _breadcrumbLabel.Text = "";
            _nodeTypeTextRadio.Checked = true;
            _nodeRefCombo.Items.Clear();
        }
        finally { _suppressEvents = false; }
        UpdateNodeTypeUI();
    }

    private string GetBreadcrumb()
    {
        if (_treeView.SelectedNode == null) return "";
        var parts = new List<string>();
        var node = _treeView.SelectedNode;
        while (node != null)
        {
            var pickNode = node.Tag as PickListNode;
            parts.Insert(0, pickNode?.Label ?? "?");
            node = node.Parent;
        }
        return string.Join(" > ", parts);
    }

    private void NodeLabelBox_TextChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedNode == null) return;
        _selectedNode.Label = _nodeLabelBox.Text;
        UpdateSelectedTreeNode();
        _breadcrumbLabel.Text = GetBreadcrumb();
    }

    private void NodeTextBox_TextChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedNode == null) return;
        _selectedNode.Text = _nodeTextBox.Text;
        UpdateSelectedTreeNode();
    }

    private void NodeTypeRadio_CheckedChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedNode == null) return;

        if (_nodeTypeBuilderRadio.Checked)
        {
            // Switching to builder reference mode - clear text and macro
            _selectedNode.Text = "";
            _nodeTextBox.Text = "";
            _selectedNode.MacroId = null;
            PopulateBuilderRefCombo();
            if (_nodeRefCombo.SelectedIndex < 0 && _nodeRefCombo.Items.Count > 0)
            {
                _nodeRefCombo.SelectedIndex = 0;
            }
        }
        else if (_nodeTypeMacroRadio.Checked)
        {
            // Switching to macro reference mode - clear text and builder
            _selectedNode.Text = "";
            _nodeTextBox.Text = "";
            _selectedNode.BuilderListId = null;
            PopulateMacroRefCombo();
            if (_nodeRefCombo.SelectedIndex < 0 && _nodeRefCombo.Items.Count > 0)
            {
                _nodeRefCombo.SelectedIndex = 0;
            }
        }
        else
        {
            // Switching to text leaf mode - clear builder and macro references
            _selectedNode.BuilderListId = null;
            _selectedNode.MacroId = null;
        }

        UpdateNodeTypeUI();
        UpdateSelectedTreeNode();
        _listBox.Invalidate();
    }

    private void NodeRefCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedNode == null) return;

        if (_nodeTypeBuilderRadio.Checked)
        {
            _selectedNode.BuilderListId = (_nodeRefCombo.SelectedItem is BuilderListItem item) ? item.Id : null;
        }
        else if (_nodeTypeMacroRadio.Checked)
        {
            _selectedNode.MacroId = (_nodeRefCombo.SelectedItem is MacroListItem item) ? item.Id : null;
        }

        UpdateSelectedTreeNode();
        _listBox.Invalidate();
    }

    private void PopulateMacroRefCombo()
    {
        _nodeRefCombo.Items.Clear();

        foreach (var macro in _config.Macros)
        {
            if (macro.Enabled)
            {
                _nodeRefCombo.Items.Add(new MacroListItem { Id = macro.Id, Name = macro.Name });
            }
        }

        // Select the currently referenced macro if any
        if (_selectedNode?.MacroId != null)
        {
            for (int i = 0; i < _nodeRefCombo.Items.Count; i++)
            {
                if (_nodeRefCombo.Items[i] is MacroListItem item && item.Id == _selectedNode.MacroId)
                {
                    _nodeRefCombo.SelectedIndex = i;
                    return;
                }
            }
        }
    }

    // Helper class for macro list dropdown items
    private class MacroListItem
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public override string ToString() => string.IsNullOrEmpty(Name) ? "(unnamed)" : Name;
    }

    private void PopulateBuilderRefCombo()
    {
        _nodeRefCombo.Items.Clear();

        // Add all enabled builder-mode pick lists (except the current one to prevent cycles)
        foreach (var list in _pickLists)
        {
            if (list.Mode == PickListMode.Builder && list.Enabled && list.Id != _selectedList?.Id)
            {
                _nodeRefCombo.Items.Add(new BuilderListItem { Id = list.Id, Name = list.Name });
            }
        }

        // Select the currently referenced builder if any
        if (_selectedNode?.BuilderListId != null)
        {
            for (int i = 0; i < _nodeRefCombo.Items.Count; i++)
            {
                if (_nodeRefCombo.Items[i] is BuilderListItem item && item.Id == _selectedNode.BuilderListId)
                {
                    _nodeRefCombo.SelectedIndex = i;
                    return;
                }
            }
        }
    }

    private void UpdateNodeTypeUI()
    {
        var isBuilderRef = _nodeTypeBuilderRadio.Checked;
        var isMacroRef = _nodeTypeMacroRadio.Checked;
        var isRef = isBuilderRef || isMacroRef;
        var isLeaf = _selectedNode != null && !_selectedNode.HasChildren;

        // Show/hide node type panel (only for leaf nodes)
        _nodeTypePanel.Visible = isLeaf;

        // Show/hide single reference combo
        _nodeRefLabel.Visible = isLeaf && isRef;
        _nodeRefCombo.Visible = isLeaf && isRef;
        _nodeRefCombo.Enabled = isLeaf && isRef;

        // Update label text based on type
        if (isBuilderRef)
        {
            _nodeRefLabel.Text = "Builder list:";
            _nodeRefLabel.ForeColor = Color.FromArgb(255, 180, 100);
        }
        else if (isMacroRef)
        {
            _nodeRefLabel.Text = "Macro:";
            _nodeRefLabel.ForeColor = Color.FromArgb(180, 150, 255);
        }

        // Text box: show for text leaves only, hide for builder/macro refs
        _prefixTextLabel.Visible = false;  // No longer used
        _textToInsertLabel.Visible = isLeaf && !isRef;
        _nodeTextBox.Visible = isLeaf && !isRef;
    }

    // Helper class for builder list dropdown items
    private class BuilderListItem
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public override string ToString() => string.IsNullOrEmpty(Name) ? "(unnamed)" : Name;
    }

    private void UpdateSelectedTreeNode()
    {
        if (_treeView.SelectedNode == null || _selectedNode == null) return;
        var idx = _treeView.SelectedNode.Parent?.Nodes.IndexOf(_treeView.SelectedNode) ?? _treeView.Nodes.IndexOf(_treeView.SelectedNode);
        var label = string.IsNullOrEmpty(_selectedNode.Label) ? "(unnamed)" : _selectedNode.Label;

        string suffix;
        Color color;

        if (_selectedNode.HasChildren)
        {
            suffix = $" [{_selectedNode.Children.Count}]";
            color = Color.FromArgb(100, 180, 255);  // Blue for branch
        }
        else if (_selectedNode.IsMacroRef)
        {
            suffix = " [M]";
            color = Color.FromArgb(180, 150, 255);  // Purple for macro ref
        }
        else if (_selectedNode.IsBuilderRef)
        {
            suffix = " [B]";
            color = Color.FromArgb(255, 180, 100);  // Orange for builder ref
        }
        else
        {
            suffix = !string.IsNullOrEmpty(_selectedNode.Text) ? " *" : "";
            color = Color.White;  // White for text leaf
        }

        _treeView.SelectedNode.Text = $"{idx + 1}. {label}{suffix}";
        _treeView.SelectedNode.ForeColor = color;
    }

    #endregion

    #region Builder Mode

    private void RefreshCategoryList()
    {
        _categoryListBox.Items.Clear();
        _selectedCategory = null;
        ClearCategoryProperties();

        if (_selectedList == null) return;

        foreach (var cat in _selectedList.Categories)
            _categoryListBox.Items.Add(cat);

        if (_categoryListBox.Items.Count > 0)
            _categoryListBox.SelectedIndex = 0;

        UpdateBuilderButtonStates();
    }

    private void CategoryListBox_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || _selectedList == null || e.Index >= _selectedList.Categories.Count) return;

        var cat = _selectedList.Categories[e.Index];
        e.DrawBackground();

        var isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        using var bgBrush = new SolidBrush(isSelected ? Color.FromArgb(50, 80, 110) : Color.FromArgb(45, 45, 45));
        e.Graphics.FillRectangle(bgBrush, e.Bounds);

        // Number
        using var numBrush = new SolidBrush(Color.FromArgb(100, 180, 255));
        using var numFont = new Font("Consolas", 10, FontStyle.Bold);
        e.Graphics.DrawString($"{e.Index + 1}.", numFont, numBrush, e.Bounds.X + 6, e.Bounds.Y + 3);

        // Name
        using var textBrush = new SolidBrush(Color.White);
        var name = string.IsNullOrEmpty(cat.Name) ? "(unnamed)" : cat.Name;
        e.Graphics.DrawString(name, e.Font!, textBrush, e.Bounds.X + 30, e.Bounds.Y + 3);

        // Option count
        using var countBrush = new SolidBrush(Color.FromArgb(110, 110, 110));
        var countText = $"({cat.Options.Count})";
        var countSize = e.Graphics.MeasureString(countText, e.Font!);
        e.Graphics.DrawString(countText, e.Font!, countBrush, e.Bounds.Right - countSize.Width - 8, e.Bounds.Y + 3);
    }

    private void CategoryListBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_categoryListBox.SelectedIndex >= 0 && _selectedList != null && _categoryListBox.SelectedIndex < _selectedList.Categories.Count)
        {
            _selectedCategory = _selectedList.Categories[_categoryListBox.SelectedIndex];
            LoadCategoryProperties();
            RefreshOptionsList();
        }
        else
        {
            _selectedCategory = null;
            ClearCategoryProperties();
            _optionsListBox.Items.Clear();
        }
        UpdateBuilderButtonStates();
    }

    private void LoadCategoryProperties()
    {
        if (_selectedCategory == null) return;
        _suppressEvents = true;
        try
        {
            _categoryNameBox.Text = _selectedCategory.Name;
            _separatorBox.Text = _selectedCategory.Separator;
            _optionsHeader.Text = $"OPTIONS IN: \"{_selectedCategory.Name}\"";
        }
        finally { _suppressEvents = false; }
    }

    private void ClearCategoryProperties()
    {
        _suppressEvents = true;
        try
        {
            _categoryNameBox.Text = "";
            _separatorBox.Text = " ";
            _optionsHeader.Text = "OPTIONS IN: (select category)";
        }
        finally { _suppressEvents = false; }
    }

    private void CategoryNameBox_TextChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedCategory == null) return;
        _selectedCategory.Name = _categoryNameBox.Text;
        _optionsHeader.Text = $"OPTIONS IN: \"{_selectedCategory.Name}\"";
        _categoryListBox.Invalidate();
    }

    private void SeparatorBox_TextChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedCategory == null) return;
        _selectedCategory.Separator = _separatorBox.Text;
    }

    private void RefreshOptionsList()
    {
        _optionsListBox.Items.Clear();
        _selectedOptionIndex = -1;
        ClearOptionProperties();

        if (_selectedCategory == null) return;

        foreach (var opt in _selectedCategory.Options)
            _optionsListBox.Items.Add(opt);

        if (_optionsListBox.Items.Count > 0)
            _optionsListBox.SelectedIndex = 0;

        UpdateBuilderButtonStates();
    }

    private string GetOptionDisplayText(BuilderOption opt)
    {
        if (opt.IsMacroRef)
        {
            var macro = _config.Macros.FirstOrDefault(m => m.Id == opt.MacroId);
            var name = macro?.Name ?? "?";
            return $"[M] {name}";
        }
        if (opt.IsTreeRef)
        {
            var tree = _config.PickLists.FirstOrDefault(p => p.Id == opt.TreeListId);
            var name = tree?.Name ?? "?";
            return $"[T] {name}";
        }
        return opt.Text;
    }

    private void OptionsListBox_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || _selectedCategory == null || e.Index >= _selectedCategory.Options.Count) return;

        var opt = _selectedCategory.Options[e.Index];
        var isTerminal = _selectedCategory.IsTerminal(e.Index);
        e.DrawBackground();

        var isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        var bgColor = isTerminal
            ? (isSelected ? Color.FromArgb(70, 60, 40) : Color.FromArgb(55, 50, 35))
            : opt.IsMacroRef ? (isSelected ? Color.FromArgb(55, 45, 75) : Color.FromArgb(45, 40, 55))
            : opt.IsTreeRef ? (isSelected ? Color.FromArgb(40, 55, 75) : Color.FromArgb(35, 45, 55))
            : (isSelected ? Color.FromArgb(50, 80, 110) : Color.FromArgb(45, 45, 45));
        using var bgBrush = new SolidBrush(bgColor);
        e.Graphics.FillRectangle(bgBrush, e.Bounds);

        // Number
        var numColor = isTerminal ? Color.FromArgb(255, 200, 100)
            : opt.IsMacroRef ? Color.FromArgb(180, 150, 255)
            : opt.IsTreeRef ? Color.FromArgb(100, 180, 255)
            : Color.FromArgb(100, 220, 150);
        using var numBrush = new SolidBrush(numColor);
        using var numFont = new Font("Consolas", 10, FontStyle.Bold);
        e.Graphics.DrawString($"{e.Index + 1}", numFont, numBrush, e.Bounds.X + 6, e.Bounds.Y + 3);

        // Option text
        using var textBrush = new SolidBrush(Color.White);
        var displayText = GetOptionDisplayText(opt);
        var text = string.IsNullOrEmpty(displayText) ? "(empty)" : displayText;
        e.Graphics.DrawString(text, e.Font!, textBrush, e.Bounds.X + 30, e.Bounds.Y + 3);

        // Indicators
        if (isTerminal)
        {
            using var endBrush = new SolidBrush(Color.FromArgb(255, 180, 100));
            e.Graphics.DrawString("[END]", e.Font!, endBrush, e.Bounds.Right - 45, e.Bounds.Y + 3);
        }
        else if (opt.IsMacroRef)
        {
            using var refBrush = new SolidBrush(Color.FromArgb(180, 150, 255));
            using var refFont = new Font("Segoe UI", 7, FontStyle.Bold);
            e.Graphics.DrawString("[M]", refFont, refBrush, e.Bounds.Right - 30, e.Bounds.Y + 5);
        }
        else if (opt.IsTreeRef)
        {
            using var refBrush = new SolidBrush(Color.FromArgb(100, 180, 255));
            using var refFont = new Font("Segoe UI", 7, FontStyle.Bold);
            e.Graphics.DrawString("[T]", refFont, refBrush, e.Bounds.Right - 25, e.Bounds.Y + 5);
        }
    }

    private void OptionsListBox_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_optionsListBox.SelectedIndex >= 0 && _selectedCategory != null && _optionsListBox.SelectedIndex < _selectedCategory.Options.Count)
        {
            _selectedOptionIndex = _optionsListBox.SelectedIndex;
            LoadOptionProperties();
        }
        else
        {
            _selectedOptionIndex = -1;
            ClearOptionProperties();
        }
        UpdateBuilderButtonStates();
    }

    private void LoadOptionProperties()
    {
        if (_selectedCategory == null || _selectedOptionIndex < 0) return;
        _suppressEvents = true;
        try
        {
            var opt = _selectedCategory.Options[_selectedOptionIndex];
            _optionTextBox.Text = opt.Text;
            _terminalCheckBox.Checked = _selectedCategory.IsTerminal(_selectedOptionIndex);

            // Set option type radio
            if (opt.IsMacroRef)
            {
                _optTypeMacroRadio.Checked = true;
                PopulateOptMacroRefCombo();
            }
            else if (opt.IsTreeRef)
            {
                _optTypeTreeRadio.Checked = true;
                PopulateOptTreeRefCombo();
            }
            else
            {
                _optTypeTextRadio.Checked = true;
            }
        }
        finally { _suppressEvents = false; }
        UpdateOptTypeUI();
    }

    private void ClearOptionProperties()
    {
        _suppressEvents = true;
        try
        {
            _optionTextBox.Text = "";
            _terminalCheckBox.Checked = false;
            _optTypeTextRadio.Checked = true;
            _optMacroRefCombo.Items.Clear();
            _optTreeRefCombo.Items.Clear();
        }
        finally { _suppressEvents = false; }
        UpdateOptTypeUI();
    }

    private void OptionTextBox_TextChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedCategory == null || _selectedOptionIndex < 0) return;
        _selectedCategory.Options[_selectedOptionIndex].Text = _optionTextBox.Text;
        _optionsListBox.Invalidate();
    }

    private void TerminalCheckBox_CheckedChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedCategory == null || _selectedOptionIndex < 0) return;

        if (_terminalCheckBox.Checked)
        {
            if (!_selectedCategory.TerminalOptions.Contains(_selectedOptionIndex))
                _selectedCategory.TerminalOptions.Add(_selectedOptionIndex);
        }
        else
        {
            _selectedCategory.TerminalOptions.Remove(_selectedOptionIndex);
        }
        _optionsListBox.Invalidate();
    }

    private void OptTypeRadio_CheckedChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedCategory == null || _selectedOptionIndex < 0) return;

        var opt = _selectedCategory.Options[_selectedOptionIndex];

        if (_optTypeMacroRadio.Checked)
        {
            opt.TreeListId = null;
            PopulateOptMacroRefCombo();
            if (_optMacroRefCombo.SelectedIndex < 0 && _optMacroRefCombo.Items.Count > 0)
                _optMacroRefCombo.SelectedIndex = 0;
        }
        else if (_optTypeTreeRadio.Checked)
        {
            opt.MacroId = null;
            PopulateOptTreeRefCombo();
            if (_optTreeRefCombo.SelectedIndex < 0 && _optTreeRefCombo.Items.Count > 0)
                _optTreeRefCombo.SelectedIndex = 0;
        }
        else
        {
            // Text mode
            opt.MacroId = null;
            opt.TreeListId = null;
        }

        UpdateOptTypeUI();
        _optionsListBox.Invalidate();
    }

    private void UpdateOptTypeUI()
    {
        var isMacro = _optTypeMacroRadio.Checked;
        var isTree = _optTypeTreeRadio.Checked;
        var hasOpt = _selectedOptionIndex >= 0;

        // Text controls
        foreach (Control c in _builderModePanel.Controls)
        {
            if (c is Label lbl && lbl.Text == "Option text:")
                lbl.Visible = hasOpt && !isMacro && !isTree;
        }
        _optionTextBox.Visible = hasOpt && !isMacro && !isTree;

        // Macro ref controls
        _optMacroRefLabel.Visible = hasOpt && isMacro;
        _optMacroRefCombo.Visible = hasOpt && isMacro;

        // Tree ref controls
        _optTreeRefLabel.Visible = hasOpt && isTree;
        _optTreeRefCombo.Visible = hasOpt && isTree;

        // Option type panel
        _optTypePanel.Visible = hasOpt;
    }

    private void PopulateOptMacroRefCombo()
    {
        _optMacroRefCombo.Items.Clear();

        foreach (var macro in _config.Macros)
        {
            if (macro.Enabled)
            {
                _optMacroRefCombo.Items.Add(new MacroListItem { Id = macro.Id, Name = macro.Name });
            }
        }

        // Select current
        if (_selectedCategory != null && _selectedOptionIndex >= 0)
        {
            var opt = _selectedCategory.Options[_selectedOptionIndex];
            if (opt.MacroId != null)
            {
                for (int i = 0; i < _optMacroRefCombo.Items.Count; i++)
                {
                    if (_optMacroRefCombo.Items[i] is MacroListItem item && item.Id == opt.MacroId)
                    {
                        _optMacroRefCombo.SelectedIndex = i;
                        return;
                    }
                }
            }
        }
    }

    private void PopulateOptTreeRefCombo()
    {
        _optTreeRefCombo.Items.Clear();

        foreach (var list in _pickLists)
        {
            if (list.Mode == PickListMode.Tree && list.Enabled && list.Id != _selectedList?.Id)
            {
                _optTreeRefCombo.Items.Add(new BuilderListItem { Id = list.Id, Name = list.Name });
            }
        }

        // Select current
        if (_selectedCategory != null && _selectedOptionIndex >= 0)
        {
            var opt = _selectedCategory.Options[_selectedOptionIndex];
            if (opt.TreeListId != null)
            {
                for (int i = 0; i < _optTreeRefCombo.Items.Count; i++)
                {
                    if (_optTreeRefCombo.Items[i] is BuilderListItem item && item.Id == opt.TreeListId)
                    {
                        _optTreeRefCombo.SelectedIndex = i;
                        return;
                    }
                }
            }
        }
    }

    private void OptMacroRefCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedCategory == null || _selectedOptionIndex < 0) return;

        var opt = _selectedCategory.Options[_selectedOptionIndex];
        if (_optMacroRefCombo.SelectedItem is MacroListItem item)
        {
            opt.MacroId = item.Id;
        }
        else
        {
            opt.MacroId = null;
        }
        _optionsListBox.Invalidate();
    }

    private void OptTreeRefCombo_SelectedIndexChanged(object? sender, EventArgs e)
    {
        if (_suppressEvents || _selectedCategory == null || _selectedOptionIndex < 0) return;

        var opt = _selectedCategory.Options[_selectedOptionIndex];
        if (_optTreeRefCombo.SelectedItem is BuilderListItem item)
        {
            opt.TreeListId = item.Id;
        }
        else
        {
            opt.TreeListId = null;
        }
        _optionsListBox.Invalidate();
    }

    private void UpdateBuilderButtonStates()
    {
        var hasList = _selectedList != null;
        var hasCat = _selectedCategory != null;
        var hasOpt = _selectedOptionIndex >= 0;

        _categoryNameBox.Enabled = hasCat;
        _separatorBox.Enabled = hasCat;
        _optionTextBox.Enabled = hasOpt;
        _terminalCheckBox.Enabled = hasOpt;

        _addCategoryBtn.Enabled = hasList;
        _removeCategoryBtn.Enabled = hasCat;
        _moveCategoryUpBtn.Enabled = hasCat && _categoryListBox.SelectedIndex > 0;
        _moveCategoryDownBtn.Enabled = hasCat && _selectedList != null && _categoryListBox.SelectedIndex < _selectedList.Categories.Count - 1;

        _addOptionBtn.Enabled = hasCat;
        _removeOptionBtn.Enabled = hasOpt;
        _moveOptionUpBtn.Enabled = hasOpt && _selectedOptionIndex > 0;
        _moveOptionDownBtn.Enabled = hasOpt && _selectedCategory != null && _selectedOptionIndex < _selectedCategory.Options.Count - 1;
    }

    private void AddCategory()
    {
        if (_selectedList == null) return;

        var newCat = new PickListCategory { Name = "New Category", Separator = " " };
        _selectedList.Categories.Add(newCat);
        RefreshCategoryList();
        _categoryListBox.SelectedIndex = _selectedList.Categories.Count - 1;
        _categoryNameBox.Focus();
        _categoryNameBox.SelectAll();
        RefreshListBox();
    }

    private void RemoveCategory()
    {
        if (_selectedList == null || _selectedCategory == null) return;

        var idx = _selectedList.Categories.IndexOf(_selectedCategory);
        _selectedList.Categories.Remove(_selectedCategory);
        RefreshCategoryList();
        if (_selectedList.Categories.Count > 0)
            _categoryListBox.SelectedIndex = Math.Min(idx, _selectedList.Categories.Count - 1);
        RefreshListBox();
    }

    private void MoveCategory(int dir)
    {
        if (_selectedList == null || _selectedCategory == null) return;

        var idx = _selectedList.Categories.IndexOf(_selectedCategory);
        var newIdx = idx + dir;
        if (newIdx < 0 || newIdx >= _selectedList.Categories.Count) return;

        _selectedList.Categories.RemoveAt(idx);
        _selectedList.Categories.Insert(newIdx, _selectedCategory);
        RefreshCategoryList();
        _categoryListBox.SelectedIndex = newIdx;
    }

    private void AddOption()
    {
        if (_selectedCategory == null) return;

        _selectedCategory.Options.Add(new BuilderOption { Text = "New option" });
        RefreshOptionsList();
        _optionsListBox.SelectedIndex = _selectedCategory.Options.Count - 1;
        _optionTextBox.Focus();
        _optionTextBox.SelectAll();
        _categoryListBox.Invalidate();
    }

    private void RemoveOption()
    {
        if (_selectedCategory == null || _selectedOptionIndex < 0) return;

        var idx = _selectedOptionIndex;
        _selectedCategory.Options.RemoveAt(idx);

        // Update terminal indices: remove this index and shift higher indices down
        _selectedCategory.TerminalOptions.Remove(idx);
        for (int i = 0; i < _selectedCategory.TerminalOptions.Count; i++)
        {
            if (_selectedCategory.TerminalOptions[i] > idx)
                _selectedCategory.TerminalOptions[i]--;
        }

        RefreshOptionsList();
        if (_selectedCategory.Options.Count > 0)
            _optionsListBox.SelectedIndex = Math.Min(idx, _selectedCategory.Options.Count - 1);
        _categoryListBox.Invalidate();
    }

    private void MoveOption(int dir)
    {
        if (_selectedCategory == null || _selectedOptionIndex < 0) return;

        var idx = _selectedOptionIndex;
        var newIdx = idx + dir;
        if (newIdx < 0 || newIdx >= _selectedCategory.Options.Count) return;

        // Move the option
        var opt = _selectedCategory.Options[idx];
        _selectedCategory.Options.RemoveAt(idx);
        _selectedCategory.Options.Insert(newIdx, opt);

        // Update terminal indices: swap the indices if needed
        var wasTerminal = _selectedCategory.TerminalOptions.Remove(idx);
        var otherWasTerminal = _selectedCategory.TerminalOptions.Remove(newIdx);

        if (wasTerminal)
            _selectedCategory.TerminalOptions.Add(newIdx);
        if (otherWasTerminal)
            _selectedCategory.TerminalOptions.Add(idx);

        RefreshOptionsList();
        _optionsListBox.SelectedIndex = newIdx;
    }

    private void ImportCategories()
    {
        if (_selectedList == null) return;

        var input = InputBox.Show(
            "Paste categories in format:\nCategory:opt1/opt2/opt3\nCategory2:opt1/opt2\n\nExample:\nSeverity:Mild/Moderate/Severe\nLocation:aortic arch/cavernous segments",
            "Import Categories",
            "",
            true);

        if (string.IsNullOrWhiteSpace(input)) return;

        var lines = input.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var imported = 0;

        foreach (var line in lines)
        {
            var colonIdx = line.IndexOf(':');
            if (colonIdx <= 0) continue;

            var name = line.Substring(0, colonIdx).Trim();
            var optionsStr = line.Substring(colonIdx + 1).Trim();
            var options = optionsStr.Split('/').Select(o => o.Trim()).Where(o => !string.IsNullOrEmpty(o))
                .Select(text => new BuilderOption { Text = text }).ToList();

            if (options.Count == 0) continue;

            _selectedList.Categories.Add(new PickListCategory
            {
                Name = name,
                Options = options,
                Separator = " "
            });
            imported++;
        }

        if (imported > 0)
        {
            RefreshCategoryList();
            RefreshListBox();
            MessageBox.Show($"Imported {imported} categories.", "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show("No valid categories found in the input.\n\nExpected format: Category:opt1/opt2/opt3", "Import Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    #endregion

    #region Button States

    private void UpdateButtonStates()
    {
        var hasList = _selectedList != null;
        var hasNode = _selectedNode != null;

        _listNameBox.Enabled = hasList;
        _modeCombo.Enabled = hasList;
        _listCriteriaRequiredBox.Enabled = hasList;
        _listCriteriaAnyOfBox.Enabled = hasList;
        _listCriteriaExcludeBox.Enabled = hasList;
        _listEnabledCheck.Enabled = hasList;

        if (_selectedList?.Mode == PickListMode.Builder)
        {
            UpdateBuilderButtonStates();
            return;
        }

        // Tree mode button states
        _addNodeBtn.Enabled = hasList;
        _addChildBtn.Enabled = hasNode;
        _cloneNodeBtn.Enabled = hasNode;
        _removeNodeBtn.Enabled = hasNode;
        _nodeLabelBox.Enabled = hasNode;

        // Leaf/builder-ref/macro-ref handling
        var isLeaf = hasNode && !_selectedNode!.HasChildren;
        var isBuilderRef = hasNode && _selectedNode!.IsBuilderRef;
        var isMacroRef = hasNode && _selectedNode!.IsMacroRef;

        // Node type panel only visible for leaf nodes (including builder/macro refs)
        _nodeTypePanel.Visible = isLeaf;
        _nodeTypePanel.Enabled = isLeaf;

        // Node ref controls (shared combo for builder and macro refs)
        // Use radio button state, not data model, since radios may change before data is saved
        var isRefRadio = _nodeTypeBuilderRadio.Checked || _nodeTypeMacroRadio.Checked;
        _nodeRefLabel.Visible = isLeaf && isRefRadio;
        _nodeRefCombo.Visible = isLeaf && isRefRadio;
        _nodeRefCombo.Enabled = isLeaf && isRefRadio;

        // Text box: show for text leaves only, hide for builder/macro refs and branches
        _textToInsertLabel.Visible = isLeaf && !isBuilderRef && !isMacroRef;
        _nodeTextBox.Visible = isLeaf && !isBuilderRef && !isMacroRef;
        _nodeTextBox.Enabled = isLeaf && !isBuilderRef && !isMacroRef;

        if (hasNode && !isLeaf && _nodeTextBox.Visible)
        {
            _suppressEvents = true;
            try { _nodeTextBox.Text = "(Branch nodes don't have text - select a leaf node)"; }
            finally { _suppressEvents = false; }
        }

        _moveUpBtn.Enabled = hasNode && CanMove(-1);
        _moveDownBtn.Enabled = hasNode && CanMove(1);
    }

    private bool CanMove(int dir)
    {
        var tn = _treeView.SelectedNode;
        if (tn == null) return false;
        var siblings = tn.Parent?.Nodes ?? _treeView.Nodes;
        var idx = siblings.IndexOf(tn);
        return idx + dir >= 0 && idx + dir < siblings.Count;
    }

    #endregion

    #region List Operations

    private void AddList()
    {
        var newList = new PickListConfig { Name = "New Pick List", Enabled = true };
        _pickLists.Add(newList);
        RefreshListBox();
        _listBox.SelectedIndex = _pickLists.Count - 1;
        _listNameBox.Focus();
        _listNameBox.SelectAll();
    }

    private void RemoveList()
    {
        if (_selectedList == null) return;
        var idx = _pickLists.IndexOf(_selectedList);
        _pickLists.Remove(_selectedList);
        RefreshListBox();
        if (_pickLists.Count > 0)
            _listBox.SelectedIndex = Math.Min(idx, _pickLists.Count - 1);
    }

    private void MoveList(int dir)
    {
        if (_selectedList == null) return;
        var idx = _pickLists.IndexOf(_selectedList);
        var newIdx = idx + dir;
        if (newIdx < 0 || newIdx >= _pickLists.Count) return;

        _pickLists.RemoveAt(idx);
        _pickLists.Insert(newIdx, _selectedList);
        RefreshListBox();
        _listBox.SelectedIndex = newIdx;
    }

    private void CloneList()
    {
        if (_selectedList == null) return;
        var clone = ClonePickList(_selectedList);
        clone.Name += " (Copy)";
        clone.Id = Guid.NewGuid().ToString("N")[..8];
        _pickLists.Add(clone);
        RefreshListBox();
        _listBox.SelectedIndex = _pickLists.Count - 1;
    }

    private void AddExampleBuilderList()
    {
        var example = new PickListConfig
        {
            Name = "Atherosclerosis (Example)",
            Mode = PickListMode.Builder,
            Enabled = true,
            CriteriaRequired = "",
            Categories = new List<PickListCategory>
            {
                new PickListCategory
                {
                    Name = "Severity",
                    Separator = " ",
                    Options = new List<BuilderOption>
                    {
                        new() { Text = "No significant atherosclerosis." },
                        new() { Text = "Trace" }, new() { Text = "Very minimal" }, new() { Text = "Minimal" }, new() { Text = "Mild" },
                        new() { Text = "Mild to moderate" }, new() { Text = "Moderate" }, new() { Text = "Moderate to severe" }, new() { Text = "Severe" }
                    },
                    TerminalOptions = new List<int> { 0 }
                },
                new PickListCategory
                {
                    Name = "Type",
                    Separator = " ",
                    Options = new List<BuilderOption>
                    {
                        new() { Text = "atherosclerotic calcifications are noted in" },
                        new() { Text = "mixed soft/hard atherosclerotic plaque is noted in" },
                        new() { Text = "predominantly hard atherosclerotic plaque is noted in" },
                        new() { Text = "predominantly soft atherosclerotic plaque is noted in" }
                    }
                },
                new PickListCategory
                {
                    Name = "Location",
                    Separator = ", producing ",
                    Options = new List<BuilderOption>
                    {
                        new() { Text = "the cavernous segments of both ICAs" },
                        new() { Text = "the supraclinoid segments of both ICAs" },
                        new() { Text = "the distal petrous/cavernous/supraclinoid segments of both ICAs" },
                        new() { Text = "both common carotid arterial bifurcations" },
                        new() { Text = "the proximal internal carotid arteries" },
                        new() { Text = "the left common carotid bifurcation/proximal left ICA" },
                        new() { Text = "the right common carotid bifurcation/proximal right ICA" },
                        new() { Text = "No significant stenosis." },
                        new() { Text = "the aortic arch and its proximal major branches" }
                    },
                    TerminalOptions = new List<int> { 7 }
                },
                new PickListCategory
                {
                    Name = "Result",
                    Separator = "",
                    Options = new List<BuilderOption>
                    {
                        new() { Text = "trace endoluminal surface irregularity." },
                        new() { Text = "minimal endoluminal surface irregularity." },
                        new() { Text = "mild endoluminal surface irregularity." },
                        new() { Text = "moderate endoluminal surface irregularity." },
                        new() { Text = "severe endoluminal surface irregularity." },
                        new() { Text = "equivalent endoluminal surface irregularity, without hemodynamically-significant stenosis." }
                    }
                }
            }
        };

        _pickLists.Add(example);
        RefreshListBox();
        _listBox.SelectedIndex = _pickLists.Count - 1;
    }

    private void AddExampleTreeList()
    {
        var example = new PickListConfig
        {
            Name = "Chest X-Ray (Example)",
            Mode = PickListMode.Tree,
            Enabled = true,
            CriteriaRequired = "",
            Nodes = new List<PickListNode>
            {
                new PickListNode
                {
                    Label = "Lungs",
                    Children = new List<PickListNode>
                    {
                        new PickListNode { Label = "Clear", Text = "The lungs are clear without focal consolidation, pleural effusion, or pneumothorax." },
                        new PickListNode
                        {
                            Label = "Opacity",
                            Children = new List<PickListNode>
                            {
                                new PickListNode { Label = "RLL opacity", Text = "There is an opacity in the right lower lobe, which may represent pneumonia or atelectasis." },
                                new PickListNode { Label = "LLL opacity", Text = "There is an opacity in the left lower lobe, which may represent pneumonia or atelectasis." },
                                new PickListNode { Label = "RUL opacity", Text = "There is an opacity in the right upper lobe, which may represent pneumonia or mass lesion." },
                                new PickListNode { Label = "Bilateral opacities", Text = "There are bilateral pulmonary opacities, which may represent multifocal pneumonia or edema." }
                            }
                        },
                        new PickListNode
                        {
                            Label = "Nodule",
                            Children = new List<PickListNode>
                            {
                                new PickListNode { Label = "Small nodule", Text = "There is a small pulmonary nodule measuring less than 6 mm. Follow-up per Fleischner guidelines may be considered." },
                                new PickListNode { Label = "Nodule >6mm", Text = "There is a pulmonary nodule measuring greater than 6 mm. CT chest is recommended for further evaluation." }
                            }
                        }
                    }
                },
                new PickListNode
                {
                    Label = "Heart",
                    Children = new List<PickListNode>
                    {
                        new PickListNode { Label = "Normal size", Text = "The cardiac silhouette is normal in size." },
                        new PickListNode { Label = "Mildly enlarged", Text = "The cardiac silhouette is mildly enlarged." },
                        new PickListNode { Label = "Moderately enlarged", Text = "The cardiac silhouette is moderately enlarged." },
                        new PickListNode { Label = "Cardiomegaly", Text = "There is cardiomegaly." }
                    }
                },
                new PickListNode
                {
                    Label = "Pleura",
                    Children = new List<PickListNode>
                    {
                        new PickListNode { Label = "No effusion", Text = "There is no pleural effusion." },
                        new PickListNode { Label = "Small right effusion", Text = "There is a small right pleural effusion." },
                        new PickListNode { Label = "Small left effusion", Text = "There is a small left pleural effusion." },
                        new PickListNode { Label = "Bilateral effusions", Text = "There are small bilateral pleural effusions." },
                        new PickListNode { Label = "Large right effusion", Text = "There is a large right pleural effusion with associated compressive atelectasis." }
                    }
                },
                new PickListNode
                {
                    Label = "Mediastinum",
                    Children = new List<PickListNode>
                    {
                        new PickListNode { Label = "Normal", Text = "The mediastinal contours are normal." },
                        new PickListNode { Label = "Widened", Text = "The mediastinum is widened. Clinical correlation is recommended." }
                    }
                },
                new PickListNode
                {
                    Label = "Bones",
                    Children = new List<PickListNode>
                    {
                        new PickListNode { Label = "No fracture", Text = "No acute osseous abnormality is identified." },
                        new PickListNode { Label = "Rib fracture", Text = "There is a rib fracture. Clinical correlation is recommended." },
                        new PickListNode { Label = "Degenerative changes", Text = "Degenerative changes of the thoracic spine are noted." }
                    }
                },
                new PickListNode
                {
                    Label = "Lines/Tubes",
                    Children = new List<PickListNode>
                    {
                        new PickListNode { Label = "ETT good position", Text = "The endotracheal tube tip is approximately 4 cm above the carina, in good position." },
                        new PickListNode { Label = "NG tube good position", Text = "The nasogastric tube tip is in the stomach, in good position." },
                        new PickListNode { Label = "Central line good position", Text = "The central venous catheter tip is in the SVC, in good position." }
                    }
                }
            }
        };

        _pickLists.Add(example);
        RefreshListBox();
        _listBox.SelectedIndex = _pickLists.Count - 1;
    }

    #endregion

    #region Node Operations (Tree Mode)

    private void AddNode(bool asChild)
    {
        if (_selectedList == null) return;

        var newNode = new PickListNode { Label = "New Item" };

        if (asChild && _selectedNode != null)
        {
            _selectedNode.Children.Add(newNode);
        }
        else if (_selectedNode != null && _treeView.SelectedNode != null)
        {
            var tn = _treeView.SelectedNode;
            if (tn.Parent == null)
            {
                var idx = _selectedList.Nodes.IndexOf(_selectedNode);
                _selectedList.Nodes.Insert(idx + 1, newNode);
            }
            else
            {
                var parent = tn.Parent.Tag as PickListNode;
                if (parent != null)
                {
                    var idx = parent.Children.IndexOf(_selectedNode);
                    parent.Children.Insert(idx + 1, newNode);
                }
            }
        }
        else
        {
            _selectedList.Nodes.Add(newNode);
        }

        RefreshTreeView();
        SelectNodeInTree(newNode);
        RefreshListBox();
        _nodeLabelBox.Focus();
        _nodeLabelBox.SelectAll();
    }

    private void RemoveNode()
    {
        if (_selectedList == null || _selectedNode == null) return;

        var tn = _treeView.SelectedNode;
        if (tn == null) return;

        if (tn.Parent == null)
            _selectedList.Nodes.Remove(_selectedNode);
        else
            (tn.Parent.Tag as PickListNode)?.Children.Remove(_selectedNode);

        RefreshTreeView();
        RefreshListBox();
    }

    private void MoveNode(int dir)
    {
        if (_selectedNode == null || _treeView.SelectedNode == null) return;

        var tn = _treeView.SelectedNode;
        var siblings = tn.Parent == null ? _selectedList!.Nodes : ((PickListNode)tn.Parent.Tag).Children;
        var idx = siblings.IndexOf(_selectedNode);
        var newIdx = idx + dir;

        if (newIdx < 0 || newIdx >= siblings.Count) return;

        siblings.RemoveAt(idx);
        siblings.Insert(newIdx, _selectedNode);

        RefreshTreeView();
        SelectNodeInTree(_selectedNode);
    }

    private void CloneNode()
    {
        if (_selectedList == null || _selectedNode == null || _treeView.SelectedNode == null) return;

        var tn = _treeView.SelectedNode;
        var siblings = tn.Parent == null ? _selectedList.Nodes : ((PickListNode)tn.Parent.Tag).Children;


        var clone = _selectedNode.Clone();
        clone.Label += " (Copy)";

        var idx = siblings.IndexOf(_selectedNode);
        siblings.Insert(idx + 1, clone);

        RefreshTreeView();
        SelectNodeInTree(clone);
        RefreshListBox();
    }

    private void SelectNodeInTree(PickListNode target)
    {
        TreeNode? Find(TreeNodeCollection nodes)
        {
            foreach (TreeNode tn in nodes)
            {
                if (tn.Tag == target) return tn;
                var found = Find(tn.Nodes);
                if (found != null) return found;
            }
            return null;
        }
        var node = Find(_treeView.Nodes);
        if (node != null) _treeView.SelectedNode = node;
    }

    #endregion

    #region Backup/Restore

    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private void BackupPickLists()
    {
        if (_pickLists.Count == 0)
        {
            MessageBox.Show("No pick lists to backup.", "Backup", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Title = "Backup Pick Lists",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            DefaultExt = "json",
            FileName = $"PickLists_Backup_{DateTime.Now:yyyy-MM-dd}.json"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        try
        {
            var json = JsonSerializer.Serialize(_pickLists, _jsonOptions);
            File.WriteAllText(dialog.FileName, json);
            MessageBox.Show($"Backed up {_pickLists.Count} pick list(s) to:\n{dialog.FileName}",
                "Backup Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save backup:\n{ex.Message}", "Backup Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void RestorePickLists()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Restore Pick Lists",
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            DefaultExt = "json"
        };

        if (dialog.ShowDialog() != DialogResult.OK)
            return;

        List<PickListConfig> imported;
        try
        {
            var json = File.ReadAllText(dialog.FileName);
            imported = JsonSerializer.Deserialize<List<PickListConfig>>(json, _jsonOptions) ?? new();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to read backup file:\n{ex.Message}", "Restore Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        if (imported.Count == 0)
        {
            MessageBox.Show("No pick lists found in the backup file.", "Restore", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Check for duplicates by name (case-insensitive)
        var existingNames = _pickLists.Select(p => p.Name.ToLowerInvariant()).ToHashSet();
        var duplicates = new List<PickListConfig>();
        var newItems = new List<PickListConfig>();

        foreach (var item in imported)
        {
            if (existingNames.Contains(item.Name.ToLowerInvariant()))
                duplicates.Add(item);
            else
                newItems.Add(item);
        }

        // Build confirmation message
        var message = $"Found {imported.Count} pick list(s) in backup.\n\n";

        if (newItems.Count > 0)
            message += $"New (will be added): {newItems.Count}\n";

        if (duplicates.Count > 0)
            message += $"Duplicates (will be skipped): {duplicates.Count}\n" +
                       $"  ({string.Join(", ", duplicates.Select(d => d.Name).Take(5))}" +
                       (duplicates.Count > 5 ? "..." : "") + ")\n";

        if (newItems.Count == 0)
        {
            MessageBox.Show(message + "\nNothing to import - all items already exist.",
                "Restore", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        message += $"\nImport {newItems.Count} new pick list(s)?";

        if (MessageBox.Show(message, "Confirm Restore", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        // Import new items with fresh IDs
        foreach (var item in newItems)
        {
            var clone = ClonePickList(item);
            clone.Id = Guid.NewGuid().ToString("N")[..8]; // New ID to avoid conflicts
            _pickLists.Add(clone);
        }

        RefreshListBox();
        if (_pickLists.Count > 0)
            _listBox.SelectedIndex = _pickLists.Count - 1;

        MessageBox.Show($"Imported {newItems.Count} pick list(s)." +
            (duplicates.Count > 0 ? $"\nSkipped {duplicates.Count} duplicate(s)." : ""),
            "Restore Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    #endregion

    private void SaveAndClose()
    {
        _config.PickLists = _pickLists;
        _config.PickListEditorWidth = Width;
        _config.PickListEditorHeight = Height;
        _config.Save();
        DialogResult = DialogResult.OK;
        Close();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);
        _config.PickListEditorWidth = Width;
        _config.PickListEditorHeight = Height;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _styleToolTip?.Dispose();
        }
        base.Dispose(disposing);
    }
}
