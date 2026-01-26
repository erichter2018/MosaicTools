using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Runtime popup for selecting pick list items.
/// Supports Tree mode (hierarchical navigation) and Builder mode (sentence construction).
/// </summary>
public class PickListPopupForm : Form
{
    private readonly Configuration _config;
    private readonly List<PickListConfig> _matchingLists;
    private readonly string? _studyDescription;
    private readonly Action<string> _onItemSelected;

    // Navigation state (shared)
    private PickListConfig? _currentList;
    private bool _showingLists = true;

    // Tree mode state
    private Stack<List<PickListNode>> _navigationStack = new();
    private Stack<string> _breadcrumbStack = new();
    private List<PickListNode> _currentNodes = new();

    // Tree mode accumulation state
    private Dictionary<string, (string label, string text)> _structuredSelections = new();  // Structured mode
    private List<string> _freeformSelections = new();  // Freeform mode
    private string? _currentTopLevelLabel;  // Track which top-level category we're in (structured)

    // Builder mode state
    private int _builderCategoryIndex = 0;
    private string _builderAccumulatedText = "";
    private List<string> _builderSelections = new(); // Track selections for undo
    private bool _pendingAutoSelect = false; // Deferred auto-select for single-option categories

    // Embedded builder state (when tree node references a builder list)
    private PickListConfig? _embeddedBuilderSourceList;  // The tree list we came from
    private Stack<List<PickListNode>>? _savedNavigationStack;  // Saved tree state
    private Stack<string>? _savedBreadcrumbStack;
    private List<PickListNode>? _savedCurrentNodes;
    private string _treeAccumulatedText = "";  // Prefix text from tree node before builder

    private TextBox _searchBox = null!;
    private ListBox _listBox = null!;
    private Label _headerLabel = null!;
    private Button _backBtn = null!;
    private Button _cancelBtn = null!;
    private Button _insertAllBtn = null!;
    private Button _clearBtn = null!;
    private Panel _navPanel = null!;
    private Label _previewLabel = null!;
    private Label _previewTitleLabel = null!;
    private Panel _previewPanel = null!;

    private int _hoveredIndex = -1;
    private bool _allowDeactivateClose = false;

    public PickListPopupForm(Configuration config, List<PickListConfig> matchingLists, string? studyDescription, Action<string> onItemSelected)
    {
        _config = config;
        _matchingLists = matchingLists;
        _studyDescription = studyDescription;
        _onItemSelected = onItemSelected;

        InitializeUI();

        // If only one list matches and skip option is enabled, go straight to it
        if (_matchingLists.Count == 1 && _config.PickListSkipSingleMatch)
        {
            _currentList = _matchingLists[0];
            _showingLists = false;

            if (_currentList.Mode == PickListMode.Builder)
            {
                _builderCategoryIndex = 0;
                _builderAccumulatedText = "";
                _builderSelections.Clear();
                _pendingAutoSelect = true;  // Defer to Shown event
            }
            else
            {
                _currentNodes = _currentList.Nodes;
            }
        }

        RefreshList();
    }

    private void InitializeUI()
    {
        Text = "Pick Lists";
        StartPosition = FormStartPosition.Manual;
        Location = new Point(_config.PickListPopupX, _config.PickListPopupY);
        FormBorderStyle = FormBorderStyle.None;
        BackColor = Color.FromArgb(35, 35, 35);
        ShowInTaskbar = false;
        TopMost = true;

        // Size depends on whether we need preview panel (will be adjusted in RefreshList)
        Size = new Size(380, 500);

        // Dark border
        var borderPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(1),
            BackColor = Color.FromArgb(80, 80, 80)
        };
        Controls.Add(borderPanel);

        var innerPanel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(35, 35, 35)
        };
        borderPanel.Controls.Add(innerPanel);

        var padding = 10;
        var y = padding;

        // Header (draggable)
        _headerLabel = new Label
        {
            Text = GetHeaderText(),
            Location = new Point(padding, y),
            Size = new Size(340, 22),
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.White,
            Cursor = Cursors.SizeAll
        };
        _headerLabel.MouseDown += Header_MouseDown;
        _headerLabel.MouseMove += Header_MouseMove;
        innerPanel.Controls.Add(_headerLabel);

        // Close button (position will be adjusted based on width)
        var closeBtn = new Label
        {
            Name = "closeBtn",
            Text = "X",
            Location = new Point(350, y),
            Size = new Size(20, 20),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = Color.Gray,
            TextAlign = ContentAlignment.MiddleCenter,
            Cursor = Cursors.Hand
        };
        closeBtn.Click += (s, e) => Close();
        closeBtn.MouseEnter += (s, e) => closeBtn.ForeColor = Color.White;
        closeBtn.MouseLeave += (s, e) => closeBtn.ForeColor = Color.Gray;
        innerPanel.Controls.Add(closeBtn);
        y += 26;

        // Navigation buttons panel
        _navPanel = new Panel
        {
            Location = new Point(padding, y),
            Size = new Size(360, 26),
            BackColor = Color.Transparent,
            Visible = false
        };
        innerPanel.Controls.Add(_navPanel);

        _backBtn = new Button
        {
            Text = "< Back",
            Location = new Point(0, 0),
            Size = new Size(70, 24),
            BackColor = Color.FromArgb(50, 70, 90),
            ForeColor = Color.FromArgb(150, 200, 255),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 8)
        };
        _backBtn.FlatAppearance.BorderColor = Color.FromArgb(70, 90, 110);
        _backBtn.Click += (s, e) => GoBack();
        _navPanel.Controls.Add(_backBtn);

        _cancelBtn = new Button
        {
            Text = "Cancel",
            Location = new Point(75, 0),
            Size = new Size(60, 24),
            BackColor = Color.FromArgb(80, 50, 50),
            ForeColor = Color.FromArgb(255, 180, 180),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 8)
        };
        _cancelBtn.FlatAppearance.BorderColor = Color.FromArgb(100, 60, 60);
        _cancelBtn.Click += (s, e) => CancelAndReset();
        _navPanel.Controls.Add(_cancelBtn);

        _clearBtn = new Button
        {
            Text = "Clear",
            Location = new Point(140, 0),
            Size = new Size(50, 24),
            BackColor = Color.FromArgb(70, 60, 50),
            ForeColor = Color.FromArgb(255, 200, 150),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 8),
            Visible = false
        };
        _clearBtn.FlatAppearance.BorderColor = Color.FromArgb(90, 75, 60);
        _clearBtn.Click += (s, e) => ClearTreeSelections();
        _navPanel.Controls.Add(_clearBtn);

        _insertAllBtn = new Button
        {
            Text = "Insert All",
            Location = new Point(195, 0),
            Size = new Size(70, 24),
            BackColor = Color.FromArgb(40, 80, 50),
            ForeColor = Color.FromArgb(150, 255, 180),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 8),
            Visible = false
        };
        _insertAllBtn.FlatAppearance.BorderColor = Color.FromArgb(60, 100, 70);
        _insertAllBtn.Click += (s, e) => InsertAllTreeSelections();
        _navPanel.Controls.Add(_insertAllBtn);

        // Search box
        _searchBox = new TextBox
        {
            Location = new Point(padding, y),
            Size = new Size(360, 24),
            BackColor = Color.FromArgb(50, 50, 50),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            Font = new Font("Segoe UI", 10)
        };
        _searchBox.TextChanged += (s, e) => RefreshList();
        _searchBox.KeyDown += SearchBox_KeyDown;
        innerPanel.Controls.Add(_searchBox);
        y += 30;

        // Preview panel (on right side for tree modes, at bottom for builder)
        _previewPanel = new Panel
        {
            Location = new Point(390, 36),
            Size = new Size(200, 450),
            BackColor = Color.FromArgb(25, 45, 35),
            Visible = false
        };
        innerPanel.Controls.Add(_previewPanel);

        _previewTitleLabel = new Label
        {
            Text = "Accumulated:",
            Location = new Point(5, 5),
            AutoSize = true,
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            ForeColor = Color.FromArgb(100, 180, 150)
        };
        _previewPanel.Controls.Add(_previewTitleLabel);

        _previewLabel = new Label
        {
            Text = "",
            Location = new Point(5, 25),
            Size = new Size(190, 420),
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(150, 255, 200)
        };
        _previewPanel.Controls.Add(_previewLabel);

        // List box
        _listBox = new ListBox
        {
            Location = new Point(padding, y),
            Size = new Size(360, Height - y - padding - 2),
            BackColor = Color.FromArgb(45, 45, 45),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.None,
            DrawMode = DrawMode.OwnerDrawFixed,
            ItemHeight = 30
        };
        _listBox.DrawItem += ListBox_DrawItem;
        _listBox.Click += (s, e) => SelectCurrentItem();
        _listBox.KeyDown += ListBox_KeyDown;
        _listBox.MouseMove += ListBox_MouseMove;
        _listBox.MouseLeave += (s, e) => { _hoveredIndex = -1; _listBox.Invalidate(); };
        innerPanel.Controls.Add(_listBox);

        // Handle clicking outside (with delay to prevent immediate close on hotkey activation)
        Deactivate += (s, e) =>
        {
            if (_allowDeactivateClose && !_config.PickListKeepOpen)
                Close();
        };

        // Focus search on show, and enable deactivate-close after a short delay
        Shown += (s, e) =>
        {
            Activate();
            _searchBox.Focus();

            // Handle deferred auto-selection for single-option categories
            if (_pendingAutoSelect)
            {
                _pendingAutoSelect = false;
                if (!AutoSelectSingleOptionCategories())
                {
                    RefreshList();
                }
            }

            var timer = new System.Windows.Forms.Timer { Interval = 200 };
            timer.Tick += (ts, te) =>
            {
                timer.Stop();
                timer.Dispose();
                _allowDeactivateClose = true;
            };
            timer.Start();
        };
    }

    private string GetHeaderText()
    {
        if (_showingLists)
        {
            var study = string.IsNullOrEmpty(_studyDescription) ? "" : $" ({TruncateString(_studyDescription, 20)})";
            return $"Pick Lists{study}";
        }

        if (_currentList?.Mode == PickListMode.Builder)
        {
            var totalCats = _currentList.Categories.Count;
            var currentStep = _builderCategoryIndex + 1;
            var prefix = _embeddedBuilderSourceList != null ? "[B] " : "";
            if (_builderCategoryIndex < totalCats)
            {
                var catName = _currentList.Categories[_builderCategoryIndex].Name;
                return $"{prefix}Step {currentStep}/{totalCats}: {catName}";
            }
            return prefix + _currentList.Name;
        }

        // Tree mode
        var styleSuffix = "";
        if (_currentList?.TreeStyle == TreePickListStyle.Structured)
        {
            styleSuffix = " [Structured]";
        }

        if (_breadcrumbStack.Count > 0)
        {
            return _breadcrumbStack.Peek() + styleSuffix;
        }

        return (_currentList?.Name ?? "Pick List") + styleSuffix;
    }

    private string GetBreadcrumbText()
    {
        if (_showingLists)
            return "";

        if (_currentList?.Mode == PickListMode.Builder)
        {
            if (_builderCategoryIndex > 0 || _builderSelections.Count > 0)
                return "< Back (Backspace) | Esc = Cancel";
            if (_matchingLists.Count > 1)
                return "< Back to list selection";
            return "";
        }

        // Tree mode
        if (_navigationStack.Count == 0)
        {
            if (_matchingLists.Count > 1)
                return "< Back to list selection";
            return "";
        }

        var parts = new List<string>();
        if (_currentList != null)
            parts.Add(_currentList.Name);
        parts.AddRange(_breadcrumbStack.Reverse().SkipLast(1));

        return "< " + string.Join(" > ", parts);
    }

    private string TruncateString(string s, int maxLen)
    {
        if (s.Length <= maxLen) return s;
        return s.Substring(0, maxLen - 3) + "...";
    }

    private void RefreshList()
    {
        _listBox.Items.Clear();
        var searchTerm = _searchBox.Text.Trim().ToUpperInvariant();

        _headerLabel.Text = GetHeaderText();

        // Determine mode states
        var isBuilderMode = !_showingLists && _currentList?.Mode == PickListMode.Builder;
        var isTreeMode = !_showingLists && _currentList?.Mode == PickListMode.Tree;
        var isStructuredTreeMode = isTreeMode && _currentList?.TreeStyle == TreePickListStyle.Structured;
        var isFreeformTreeMode = isTreeMode && _currentList?.TreeStyle == TreePickListStyle.Freeform;

        // Check if tree mode has accumulated selections
        var hasTreeSelections = (isStructuredTreeMode && _structuredSelections.Count > 0) ||
                                (isFreeformTreeMode && _freeformSelections.Count > 0);

        // Determine if we should show navigation buttons
        var showNav = ShouldShowNavigation() || hasTreeSelections;
        _navPanel.Visible = showNav;

        // Update button states
        if (showNav)
        {
            var canGoBack = CanGoBack();
            _backBtn.Visible = canGoBack;
            _cancelBtn.Visible = true;

            // Show tree mode buttons when there are selections
            _clearBtn.Visible = hasTreeSelections;
            _insertAllBtn.Visible = hasTreeSelections;
        }
        else
        {
            _clearBtn.Visible = false;
            _insertAllBtn.Visible = false;
        }

        // Show preview panel for tree modes (both Freeform and Structured) or Builder mode
        var showPreview = isBuilderMode || isTreeMode;
        _previewPanel.Visible = showPreview;

        if (isBuilderMode)
        {
            _previewTitleLabel.Text = "Building:";
            var displayText = string.IsNullOrEmpty(_builderAccumulatedText) ? "(empty)" : $"\"{_builderAccumulatedText}\"";
            _previewLabel.Text = displayText;
        }
        else if (isStructuredTreeMode)
        {
            _previewTitleLabel.Text = "Accumulated:";
            UpdateStructuredPreview();
        }
        else if (isFreeformTreeMode)
        {
            _previewTitleLabel.Text = "Accumulated:";
            UpdateFreeformPreview();
        }

        AdjustLayout(showNav, showPreview, isBuilderMode);

        if (_showingLists)
        {
            // Show list selector
            foreach (var list in _matchingLists)
            {
                if (string.IsNullOrEmpty(searchTerm) || list.Name.ToUpperInvariant().Contains(searchTerm))
                {
                    _listBox.Items.Add(new ListEntry { List = list });
                }
            }
        }
        else if (_currentList?.Mode == PickListMode.Builder)
        {
            // Show current category options
            if (_builderCategoryIndex < _currentList.Categories.Count)
            {
                var category = _currentList.Categories[_builderCategoryIndex];
                for (int i = 0; i < category.Options.Count && i < 9; i++)
                {
                    var opt = category.Options[i];
                    if (string.IsNullOrEmpty(searchTerm) || opt.ToUpperInvariant().Contains(searchTerm))
                    {
                        _listBox.Items.Add(new BuilderOptionEntry { Index = i, Text = opt });
                    }
                }
            }
        }
        else
        {
            // Show current nodes (Tree mode)
            for (int i = 0; i < _currentNodes.Count && i < 9; i++)
            {
                var node = _currentNodes[i];
                if (string.IsNullOrEmpty(searchTerm) ||
                    node.Label.ToUpperInvariant().Contains(searchTerm) ||
                    node.Text.ToUpperInvariant().Contains(searchTerm))
                {
                    _listBox.Items.Add(new NodeEntry { Index = i, Node = node });
                }
            }
        }

        if (_listBox.Items.Count > 0)
            _listBox.SelectedIndex = 0;
    }

    private void AdjustLayout(bool showNav, bool showPreview, bool isBuilderMode)
    {
        var padding = 10;
        var listWidth = 360;
        var previewWidth = 200;

        // Adjust window width based on whether preview is shown (and not builder mode)
        // Builder mode uses bottom preview, tree modes use side preview
        if (showPreview && !isBuilderMode)
        {
            // Wide layout with preview on right
            Size = new Size(listWidth + previewWidth + padding * 3 + 2, 500);
            _previewPanel.Location = new Point(listWidth + padding * 2, 36);
            _previewPanel.Size = new Size(previewWidth, Height - 50);
            _previewLabel.Size = new Size(previewWidth - 10, Height - 80);
        }
        else if (showPreview && isBuilderMode)
        {
            // Narrow layout with preview at bottom (for builder mode)
            Size = new Size(listWidth + padding * 2 + 2, 500);
            _previewPanel.Location = new Point(padding, Height - 70);
            _previewPanel.Size = new Size(listWidth, 55);
            _previewLabel.Size = new Size(listWidth - 10, 30);
        }
        else
        {
            // Narrow layout without preview
            Size = new Size(listWidth + padding * 2 + 2, 500);
        }

        // Position close button
        foreach (Control c in Controls)
        {
            if (c is Panel borderPanel)
            {
                foreach (Control inner in borderPanel.Controls)
                {
                    if (inner is Panel innerPanel)
                    {
                        foreach (Control ctrl in innerPanel.Controls)
                        {
                            if (ctrl.Name == "closeBtn")
                            {
                                ctrl.Location = new Point(Width - 30, 10);
                            }
                        }
                    }
                }
            }
        }

        _searchBox.Top = showNav ? _navPanel.Bottom + 5 : 36;
        _searchBox.Width = listWidth;
        _listBox.Top = _searchBox.Bottom + 6;
        _listBox.Width = listWidth;

        var bottomMargin = (showPreview && isBuilderMode) ? 75 : 12;
        _listBox.Height = Height - _listBox.Top - bottomMargin;
    }

    // Entry classes
    private class ListEntry
    {
        public PickListConfig List { get; set; } = null!;
    }

    private class NodeEntry
    {
        public int Index { get; set; }
        public PickListNode Node { get; set; } = null!;
    }

    private class BuilderOptionEntry
    {
        public int Index { get; set; }
        public string Text { get; set; } = "";
    }

    private void ListBox_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= _listBox.Items.Count) return;

        e.DrawBackground();

        var isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        var isHovered = e.Index == _hoveredIndex;
        var bgColor = isSelected ? Color.FromArgb(50, 90, 130) :
                      isHovered ? Color.FromArgb(55, 55, 55) :
                      Color.FromArgb(45, 45, 45);

        using var bgBrush = new SolidBrush(bgColor);
        e.Graphics.FillRectangle(bgBrush, e.Bounds);

        var obj = _listBox.Items[e.Index];

        if (obj is ListEntry listEntry)
        {
            // Drawing a pick list
            var modeIndicator = listEntry.List.Mode == PickListMode.Builder ? "[B] " : "[T] ";
            var modeColor = listEntry.List.Mode == PickListMode.Builder ? Color.FromArgb(255, 180, 100) : Color.FromArgb(100, 180, 255);

            using var arrowBrush = new SolidBrush(Color.FromArgb(120, 120, 120));
            using var arrowFont = new Font("Segoe UI", 10, FontStyle.Bold);
            e.Graphics.DrawString(">", arrowFont, arrowBrush, e.Bounds.X + 8, e.Bounds.Y + 5);

            using var textBrush = new SolidBrush(modeColor);
            e.Graphics.DrawString(modeIndicator + listEntry.List.Name, e.Font!, textBrush, e.Bounds.X + 28, e.Bounds.Y + 6);

            // Item count
            var count = listEntry.List.Mode == PickListMode.Builder
                ? listEntry.List.Categories.Count
                : listEntry.List.Nodes.Count + listEntry.List.Nodes.Sum(n => n.CountDescendants());
            var countText = $"({count})";
            var countSize = e.Graphics.MeasureString(countText, e.Font!);
            using var countBrush = new SolidBrush(Color.FromArgb(100, 100, 100));
            e.Graphics.DrawString(countText, e.Font!, countBrush, e.Bounds.Right - countSize.Width - 10, e.Bounds.Y + 6);
        }
        else if (obj is NodeEntry nodeEntry)
        {
            var node = nodeEntry.Node;

            // Determine node type and color
            Color numColor;
            if (node.IsBuilderRef)
                numColor = Color.FromArgb(255, 180, 100);  // Orange for builder-ref
            else if (node.HasChildren)
                numColor = Color.FromArgb(100, 180, 255);  // Blue for branch
            else
                numColor = Color.FromArgb(100, 220, 150);  // Green for leaf

            // Number
            var numText = $"{nodeEntry.Index + 1}";
            using var numBrush = new SolidBrush(numColor);
            using var numFont = new Font("Consolas", 11, FontStyle.Bold);
            e.Graphics.DrawString(numText, numFont, numBrush, e.Bounds.X + 8, e.Bounds.Y + 5);

            // Label
            using var textBrush = new SolidBrush(Color.White);
            e.Graphics.DrawString(node.Label, e.Font!, textBrush, e.Bounds.X + 32, e.Bounds.Y + 6);

            // Indicator based on node type
            if (node.IsBuilderRef)
            {
                // Builder reference - show [B] indicator
                using var builderBrush = new SolidBrush(Color.FromArgb(255, 180, 100));
                using var builderFont = new Font("Segoe UI", 7, FontStyle.Bold);
                e.Graphics.DrawString("[B]", builderFont, builderBrush, e.Bounds.Right - 30, e.Bounds.Y + 8);
            }
            else if (node.HasChildren)
            {
                var childText = $"({node.Children.Count})";
                var childSize = e.Graphics.MeasureString(childText, e.Font!);
                using var childBrush = new SolidBrush(Color.FromArgb(100, 100, 100));
                e.Graphics.DrawString(childText, e.Font!, childBrush, e.Bounds.Right - childSize.Width - 30, e.Bounds.Y + 6);

                using var arrowBrush = new SolidBrush(Color.FromArgb(100, 180, 255));
                using var arrowFont = new Font("Segoe UI", 10, FontStyle.Bold);
                e.Graphics.DrawString(">", arrowFont, arrowBrush, e.Bounds.Right - 20, e.Bounds.Y + 5);
            }
            else if (!string.IsNullOrEmpty(node.Text))
            {
                using var checkBrush = new SolidBrush(Color.FromArgb(100, 220, 150));
                using var checkFont = new Font("Segoe UI", 9);
                e.Graphics.DrawString("*", checkFont, checkBrush, e.Bounds.Right - 18, e.Bounds.Y + 4);
            }
        }
        else if (obj is BuilderOptionEntry optEntry)
        {
            // Check if terminal
            var isTerminal = _currentList != null &&
                             _builderCategoryIndex < _currentList.Categories.Count &&
                             _currentList.Categories[_builderCategoryIndex].IsTerminal(optEntry.Index);

            // Number (different color for terminal)
            var numText = $"{optEntry.Index + 1}";
            var numColor = isTerminal ? Color.FromArgb(255, 200, 100) : Color.FromArgb(100, 220, 150);
            using var numBrush = new SolidBrush(numColor);
            using var numFont = new Font("Consolas", 11, FontStyle.Bold);
            e.Graphics.DrawString(numText, numFont, numBrush, e.Bounds.X + 8, e.Bounds.Y + 5);

            // Option text
            using var textBrush = new SolidBrush(Color.White);
            var text = string.IsNullOrEmpty(optEntry.Text) ? "(empty)" : optEntry.Text;
            e.Graphics.DrawString(text, e.Font!, textBrush, e.Bounds.X + 32, e.Bounds.Y + 6);

            // Terminal indicator
            if (isTerminal)
            {
                using var endBrush = new SolidBrush(Color.FromArgb(255, 180, 100));
                using var endFont = new Font("Segoe UI", 7);
                e.Graphics.DrawString("[END]", endFont, endBrush, e.Bounds.Right - 40, e.Bounds.Y + 8);
            }
        }
    }

    private void SelectCurrentItem()
    {
        if (_listBox.SelectedIndex < 0) return;

        var obj = _listBox.SelectedItem;

        if (obj is ListEntry listEntry)
        {
            // Select a list
            _currentList = listEntry.List;
            _showingLists = false;

            if (_currentList.Mode == PickListMode.Builder)
            {
                _builderCategoryIndex = 0;
                _builderAccumulatedText = "";
                _builderSelections.Clear();

                _searchBox.Clear();
                if (!AutoSelectSingleOptionCategories())
                {
                    RefreshList();
                }
                return;  // Early return to prevent duplicate refresh
            }
            else
            {
                _currentNodes = _currentList.Nodes;
                _navigationStack.Clear();
                _breadcrumbStack.Clear();
                _structuredSelections.Clear();
                _freeformSelections.Clear();
                _currentTopLevelLabel = null;
            }

            _searchBox.Clear();
            RefreshList();
        }
        else if (obj is NodeEntry nodeEntry)
        {
            var node = nodeEntry.Node;

            if (node.IsBuilderRef)
            {
                // Builder reference node - switch to embedded builder mode
                var builderList = _config.PickLists.FirstOrDefault(p => p.Id == node.BuilderListId);
                if (builderList == null || !builderList.Enabled || builderList.Mode != PickListMode.Builder)
                {
                    // Invalid or disabled builder reference - show error
                    System.Windows.Forms.MessageBox.Show(
                        "The referenced builder list is not available (deleted, disabled, or not a builder).",
                        "Builder Not Found",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                // Save current tree state for potential cancel/return
                _embeddedBuilderSourceList = _currentList;
                _savedNavigationStack = new Stack<List<PickListNode>>(_navigationStack.Reverse());
                _savedBreadcrumbStack = new Stack<string>(_breadcrumbStack.Reverse());
                _savedCurrentNodes = _currentNodes;
                _treeAccumulatedText = "";  // Builder refs don't use prefix text

                // Switch to the builder list
                _currentList = builderList;
                _builderCategoryIndex = 0;
                _builderAccumulatedText = "";
                _builderSelections.Clear();

                _searchBox.Clear();
                if (!AutoSelectSingleOptionCategories())
                {
                    RefreshList();
                }
                return;  // Early return
            }
            else if (node.HasChildren)
            {
                // Drill down into children
                // Track top-level label for structured mode
                if (_navigationStack.Count == 0 && _currentList?.TreeStyle == TreePickListStyle.Structured)
                {
                    _currentTopLevelLabel = node.Label;
                }

                _navigationStack.Push(_currentNodes);
                _breadcrumbStack.Push(node.Label);
                _currentNodes = node.Children;
                _searchBox.Clear();
                RefreshList();
            }
            else if (!string.IsNullOrEmpty(node.Text))
            {
                // Leaf node selected - behavior depends on tree style
                if (_currentList?.TreeStyle == TreePickListStyle.Structured)
                {
                    // STRUCTURED MODE: Accumulate selection for the top-level category
                    var topLevelLabel = GetCurrentTopLevelLabel();
                    if (topLevelLabel != null)
                    {
                        // Store/replace selection for this category (use label as key)
                        _structuredSelections[topLevelLabel] = (topLevelLabel, node.Text);
                    }

                    // Go back to root level of this pick list (to select another category)
                    _navigationStack.Clear();
                    _breadcrumbStack.Clear();
                    _currentNodes = _currentList?.Nodes ?? new();
                    _currentTopLevelLabel = null;
                    _searchBox.Clear();
                    RefreshList();
                }
                else
                {
                    // FREEFORM MODE: Accumulate selection
                    _freeformSelections.Add(node.Text);

                    // Go back to root level of this pick list (to select more items)
                    _navigationStack.Clear();
                    _breadcrumbStack.Clear();
                    _currentNodes = _currentList?.Nodes ?? new();
                    _searchBox.Clear();
                    RefreshList();
                }
            }
        }
        else if (obj is BuilderOptionEntry optEntry)
        {
            if (_currentList == null || _builderCategoryIndex >= _currentList.Categories.Count) return;

            var category = _currentList.Categories[_builderCategoryIndex];
            var isTerminal = category.IsTerminal(optEntry.Index);

            // Add selection to accumulated text
            _builderSelections.Add(optEntry.Text);

            if (isTerminal)
            {
                // Terminal option - complete sentence immediately without separator
                _builderAccumulatedText += optEntry.Text;
                CompleteBuilderAndInsert();
            }
            else
            {
                // Normal option - add separator and continue
                _builderAccumulatedText += optEntry.Text + category.Separator;
                _builderCategoryIndex++;

                // Auto-select any following single-option categories
                if (!AutoSelectSingleOptionCategories())
                {
                    // Check if we've completed all categories
                    if (_builderCategoryIndex >= _currentList.Categories.Count)
                    {
                        // Trim trailing separator from last category
                        var lastSep = _currentList.Categories.Last().Separator;
                        if (_builderAccumulatedText.EndsWith(lastSep))
                        {
                            _builderAccumulatedText = _builderAccumulatedText.Substring(0, _builderAccumulatedText.Length - lastSep.Length);
                        }

                        CompleteBuilderAndInsert();
                    }
                }
            }

            _searchBox.Clear();
            RefreshList();
        }
    }

    /// <summary>
    /// Complete the builder and insert text. Handles both standalone and embedded builder modes.
    /// </summary>
    private void CompleteBuilderAndInsert()
    {
        SavePosition();
        _onItemSelected(_builderAccumulatedText);

        // Check if we were in embedded builder mode (came from a tree node)
        if (_embeddedBuilderSourceList != null)
        {
            // Return to tree mode - reset to root of the source tree list
            _currentList = _embeddedBuilderSourceList;
            _navigationStack.Clear();
            _breadcrumbStack.Clear();
            _currentNodes = _currentList.Nodes;

            // Clear embedded state
            _embeddedBuilderSourceList = null;
            _savedNavigationStack = null;
            _savedBreadcrumbStack = null;
            _savedCurrentNodes = null;
            _treeAccumulatedText = "";
        }

        // Reset builder state for next use
        _builderCategoryIndex = 0;
        _builderAccumulatedText = "";
        _builderSelections.Clear();
    }

    /// <summary>
    /// Auto-selects categories that have only one option, advancing through them
    /// without user interaction. Handles terminal options and completion.
    /// </summary>
    /// <returns>True if builder was completed (all categories done or terminal hit)</returns>
    private bool AutoSelectSingleOptionCategories()
    {
        if (_currentList?.Mode != PickListMode.Builder) return false;

        while (_builderCategoryIndex < _currentList.Categories.Count)
        {
            var category = _currentList.Categories[_builderCategoryIndex];

            // Only auto-select if exactly 1 option
            if (category.Options.Count != 1) break;

            var singleOption = category.Options[0];
            var isTerminal = category.IsTerminal(0);

            // Track for undo
            _builderSelections.Add(singleOption);

            if (isTerminal)
            {
                _builderAccumulatedText += singleOption;
                CompleteBuilderAndInsert();
                return true;
            }

            _builderAccumulatedText += singleOption + category.Separator;
            _builderCategoryIndex++;
        }

        // Check if completed all categories
        if (_builderCategoryIndex >= _currentList.Categories.Count)
        {
            var lastSep = _currentList.Categories.Last().Separator;
            if (_builderAccumulatedText.EndsWith(lastSep))
            {
                _builderAccumulatedText = _builderAccumulatedText.Substring(
                    0, _builderAccumulatedText.Length - lastSep.Length);
            }
            CompleteBuilderAndInsert();
            return true;
        }

        return false;
    }

    /// <summary>
    /// Update the preview panel for structured tree mode to show accumulated selections.
    /// </summary>
    private void UpdateStructuredPreview()
    {
        if (_structuredSelections.Count == 0)
        {
            _previewLabel.Text = "(no selections)\n\nSelect items from\ncategories, then\nclick Insert All.";
            return;
        }

        // Build preview text matching the output format
        var lines = new List<string>();
        foreach (var kvp in _structuredSelections)
        {
            // Truncate long text for preview
            var text = kvp.Value.text;
            if (text.Length > 50)
                text = text.Substring(0, 47) + "...";

            var label = kvp.Value.label.ToUpperInvariant();
            if (_currentList?.StructuredTextPlacement == StructuredTextPlacement.BelowHeading)
                lines.Add($"{label}:\n  {text}");
            else
                lines.Add($"{label}: {text}");
        }

        var separator = (_currentList?.StructuredBlankLines == true) ? "\n\n" : "\n";
        _previewLabel.Text = string.Join(separator, lines);
    }

    /// <summary>
    /// Update the preview panel for freeform tree mode to show accumulated selections.
    /// </summary>
    private void UpdateFreeformPreview()
    {
        if (_freeformSelections.Count == 0)
        {
            _previewLabel.Text = "(no selections)\n\nSelect items to\naccumulate, then\nclick Insert All.";
            return;
        }

        // Join with space, showing accumulated text
        var accumulated = string.Join(" ", _freeformSelections);
        _previewLabel.Text = accumulated;
    }

    /// <summary>
    /// Clear all accumulated tree selections (both freeform and structured).
    /// </summary>
    private void ClearTreeSelections()
    {
        _structuredSelections.Clear();
        _freeformSelections.Clear();
        _currentTopLevelLabel = null;
        RefreshList();
    }

    /// <summary>
    /// Insert all accumulated tree selections as formatted text.
    /// </summary>
    private void InsertAllTreeSelections()
    {
        string formattedText;

        if (_currentList?.TreeStyle == TreePickListStyle.Structured)
        {
            if (_structuredSelections.Count == 0) return;

            // Format based on configuration options
            var lines = new List<string>();
            foreach (var kvp in _structuredSelections)
            {
                var label = kvp.Value.label.ToUpperInvariant();
                var text = kvp.Value.text;

                if (_currentList.StructuredTextPlacement == StructuredTextPlacement.BelowHeading)
                    lines.Add($"{label}:\n  {text}");
                else
                    lines.Add($"{label}: {text}");
            }

            var separator = _currentList.StructuredBlankLines ? "\n\n" : "\n";
            formattedText = string.Join(separator, lines);
        }
        else
        {
            // Freeform mode
            if (_freeformSelections.Count == 0) return;
            formattedText = string.Join(" ", _freeformSelections);
        }

        SavePosition();
        _onItemSelected(formattedText);

        // Reset tree state but stay in the same list
        _structuredSelections.Clear();
        _freeformSelections.Clear();
        _currentTopLevelLabel = null;
        _navigationStack.Clear();
        _breadcrumbStack.Clear();
        _currentNodes = _currentList?.Nodes ?? new();
        _searchBox.Clear();
        RefreshList();
    }

    /// <summary>
    /// Get the top-level parent label for the current navigation path.
    /// </summary>
    private string? GetCurrentTopLevelLabel()
    {
        if (_breadcrumbStack.Count > 0)
        {
            // The first item in the breadcrumb is the top-level category
            return _breadcrumbStack.Reverse().First();
        }
        return _currentTopLevelLabel;
    }

    private void GoBack()
    {
        if (_currentList?.Mode == PickListMode.Builder)
        {
            // Builder mode: undo last selection
            if (_builderSelections.Count > 0)
            {
                var lastSelection = _builderSelections[_builderSelections.Count - 1];
                _builderSelections.RemoveAt(_builderSelections.Count - 1);
                _builderCategoryIndex--;

                // Remove the last selection and its separator from accumulated text
                if (_builderCategoryIndex >= 0 && _builderCategoryIndex < _currentList.Categories.Count)
                {
                    var prevCategory = _currentList.Categories[_builderCategoryIndex];
                    var toRemove = lastSelection + prevCategory.Separator;
                    if (_builderAccumulatedText.EndsWith(toRemove))
                    {
                        _builderAccumulatedText = _builderAccumulatedText.Substring(0, _builderAccumulatedText.Length - toRemove.Length);
                    }
                }

                _searchBox.Clear();
                RefreshList();
            }
            else if (_embeddedBuilderSourceList != null)
            {
                // In embedded builder with no selections - return to tree at saved position
                ReturnToTreeFromEmbeddedBuilder();
            }
            else if (_matchingLists.Count > 1)
            {
                // Go back to list selector
                _showingLists = true;
                _currentList = null;
                _searchBox.Clear();
                RefreshList();
            }
        }
        else
        {
            // Tree mode
            if (_navigationStack.Count > 0)
            {
                _currentNodes = _navigationStack.Pop();
                _breadcrumbStack.Pop();

                // Clear top-level label tracking when back at root in structured mode
                if (_navigationStack.Count == 0)
                {
                    _currentTopLevelLabel = null;
                }

                _searchBox.Clear();
                RefreshList();
            }
            else if (!_showingLists && _matchingLists.Count > 1)
            {
                _showingLists = true;
                _currentList = null;
                _currentNodes = new();
                _structuredSelections.Clear();
                _freeformSelections.Clear();
                _currentTopLevelLabel = null;
                _searchBox.Clear();
                RefreshList();
            }
        }
    }

    /// <summary>
    /// Return to the tree list from embedded builder mode, restoring saved position.
    /// </summary>
    private void ReturnToTreeFromEmbeddedBuilder()
    {
        if (_embeddedBuilderSourceList == null) return;

        // Restore tree state
        _currentList = _embeddedBuilderSourceList;
        if (_savedNavigationStack != null)
            _navigationStack = new Stack<List<PickListNode>>(_savedNavigationStack.Reverse());
        else
            _navigationStack.Clear();

        if (_savedBreadcrumbStack != null)
            _breadcrumbStack = new Stack<string>(_savedBreadcrumbStack.Reverse());
        else
            _breadcrumbStack.Clear();

        _currentNodes = _savedCurrentNodes ?? _currentList.Nodes;

        // Clear embedded state
        _embeddedBuilderSourceList = null;
        _savedNavigationStack = null;
        _savedBreadcrumbStack = null;
        _savedCurrentNodes = null;
        _treeAccumulatedText = "";

        // Reset builder state
        _builderCategoryIndex = 0;
        _builderAccumulatedText = "";
        _builderSelections.Clear();

        _searchBox.Clear();
        RefreshList();
    }

    private void CancelBuilder()
    {
        // If in embedded builder mode, return to tree
        if (_embeddedBuilderSourceList != null)
        {
            ReturnToTreeFromEmbeddedBuilder();
            return;
        }

        // Scrap current sentence and start over
        _builderCategoryIndex = 0;
        _builderAccumulatedText = "";
        _builderSelections.Clear();
        _searchBox.Clear();
        RefreshList();
    }

    private void CancelAndReset()
    {
        // If in embedded builder mode, return to tree first
        if (_embeddedBuilderSourceList != null)
        {
            ReturnToTreeFromEmbeddedBuilder();
            return;
        }

        // Full reset - go back to list selection or close
        if (_matchingLists.Count > 1)
        {
            _showingLists = true;
            _currentList = null;
            _currentNodes = new();
            _navigationStack.Clear();
            _breadcrumbStack.Clear();
            _builderCategoryIndex = 0;
            _builderAccumulatedText = "";
            _builderSelections.Clear();
            _structuredSelections.Clear();
            _freeformSelections.Clear();
            _currentTopLevelLabel = null;
            _searchBox.Clear();
            RefreshList();
        }
        else
        {
            // Only one list - reset to beginning of that list
            if (_currentList?.Mode == PickListMode.Builder)
            {
                CancelBuilder();
            }
            else
            {
                _navigationStack.Clear();
                _breadcrumbStack.Clear();
                _currentNodes = _currentList?.Nodes ?? new();
                _structuredSelections.Clear();
                _freeformSelections.Clear();
                _currentTopLevelLabel = null;
                _searchBox.Clear();
                RefreshList();
            }
        }
    }

    private bool ShouldShowNavigation()
    {
        // Show nav when we're inside a list (not at list selection)
        if (_showingLists) return false;

        // In Builder mode, show when we have selections, are embedded, or multiple lists
        if (_currentList?.Mode == PickListMode.Builder)
            return _builderSelections.Count > 0 || _embeddedBuilderSourceList != null || _matchingLists.Count > 1;

        // In Tree mode
        if (_currentList?.Mode == PickListMode.Tree)
        {
            // Show nav when we have accumulated selections
            if (_currentList.TreeStyle == TreePickListStyle.Structured && _structuredSelections.Count > 0)
                return true;
            if (_currentList.TreeStyle == TreePickListStyle.Freeform && _freeformSelections.Count > 0)
                return true;

            // Show when we're drilled down or have multiple lists
            return _navigationStack.Count > 0 || _matchingLists.Count > 1;
        }

        return _navigationStack.Count > 0 || _matchingLists.Count > 1;
    }

    private bool CanGoBack()
    {
        // Can go back one step (not full cancel)
        if (_currentList?.Mode == PickListMode.Builder)
            return _builderSelections.Count > 0 || _embeddedBuilderSourceList != null || _matchingLists.Count > 1;

        return _navigationStack.Count > 0 || _matchingLists.Count > 1;
    }

    private void SelectByNumber(int num)
    {
        // Find item with matching index
        for (int i = 0; i < _listBox.Items.Count; i++)
        {
            var item = _listBox.Items[i];
            if ((item is NodeEntry nodeEntry && nodeEntry.Index == num) ||
                (item is BuilderOptionEntry optEntry && optEntry.Index == num))
            {
                _listBox.SelectedIndex = i;
                SelectCurrentItem();
                return;
            }
        }
    }

    private void ListBox_KeyDown(object? sender, KeyEventArgs e)
    {
        // Handle number keys 1-9
        if (e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D9)
        {
            SelectByNumber(e.KeyCode - Keys.D1);
            e.Handled = true;
            return;
        }
        if (e.KeyCode >= Keys.NumPad1 && e.KeyCode <= Keys.NumPad9)
        {
            SelectByNumber(e.KeyCode - Keys.NumPad1);
            e.Handled = true;
            return;
        }

        if (e.KeyCode == Keys.Enter)
        {
            SelectCurrentItem();
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Back)
        {
            GoBack();
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Escape)
        {
            HandleEscape();
            e.Handled = true;
        }
    }

    private void SearchBox_KeyDown(object? sender, KeyEventArgs e)
    {
        // Handle number keys when search is empty
        if (string.IsNullOrEmpty(_searchBox.Text))
        {
            if (e.KeyCode >= Keys.D1 && e.KeyCode <= Keys.D9)
            {
                SelectByNumber(e.KeyCode - Keys.D1);
                e.Handled = true;
                return;
            }
            if (e.KeyCode >= Keys.NumPad1 && e.KeyCode <= Keys.NumPad9)
            {
                SelectByNumber(e.KeyCode - Keys.NumPad1);
                e.Handled = true;
                return;
            }
        }

        if (e.KeyCode == Keys.Down)
        {
            _listBox.Focus();
            if (_listBox.Items.Count > 0 && _listBox.SelectedIndex < 0)
                _listBox.SelectedIndex = 0;
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Up)
        {
            _listBox.Focus();
            if (_listBox.Items.Count > 0)
                _listBox.SelectedIndex = _listBox.Items.Count - 1;
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Enter)
        {
            if (_listBox.Items.Count > 0)
            {
                if (_listBox.SelectedIndex < 0)
                    _listBox.SelectedIndex = 0;
                SelectCurrentItem();
            }
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Back && string.IsNullOrEmpty(_searchBox.Text))
        {
            GoBack();
            e.Handled = true;
        }
        else if (e.KeyCode == Keys.Escape)
        {
            HandleEscape();
            e.Handled = true;
        }
    }

    private void HandleEscape()
    {
        // Escape always does a full cancel/reset (not step-by-step like Backspace)
        if (!_showingLists)
        {
            CancelAndReset();
        }
        else
        {
            // Already at list selection - close the popup
            SavePosition();
            Close();
        }
    }

    private void ListBox_MouseMove(object? sender, MouseEventArgs e)
    {
        var index = _listBox.IndexFromPoint(e.Location);
        if (index != _hoveredIndex)
        {
            _hoveredIndex = index;
            _listBox.Invalidate();
        }
    }

    // Dragging support
    private Point _dragOffset;
    private bool _dragging;

    private void Header_MouseDown(object? sender, MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left)
        {
            _dragging = true;
            _dragOffset = e.Location;
        }
    }

    private void Header_MouseMove(object? sender, MouseEventArgs e)
    {
        if (_dragging)
        {
            var newLocation = PointToScreen(e.Location);
            Location = new Point(newLocation.X - _dragOffset.X, newLocation.Y - _dragOffset.Y);
        }
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        base.OnMouseUp(e);
        _dragging = false;
    }

    private void SavePosition()
    {
        _config.PickListPopupX = Location.X;
        _config.PickListPopupY = Location.Y;
        _config.Save();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);
        SavePosition();
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape)
        {
            HandleEscape();
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }
}
