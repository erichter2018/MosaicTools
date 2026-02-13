using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MosaicTools.Services;

namespace MosaicTools.UI;

/// <summary>
/// Popup form showing the list of critical studies where notes were placed this session.
/// </summary>
public class CriticalStudiesPopup : Form
{
    private readonly IReadOnlyList<CriticalStudyEntry> _studies;
    private readonly AutomationService? _automationService;
    private readonly Action<CriticalStudyEntry>? _onRemoveStudy;
    private readonly ListBox _listBox;
    private bool _allowDeactivateClose = false;
    private int _hoveredIndex = -1;
    private bool _hoveringDelete = false;  // true when mouse is over the X button area

    // Delete button hit area: 24x24 in top-right of each row
    private const int DeleteBtnSize = 24;
    private const int DeleteBtnMargin = 4;

    public CriticalStudiesPopup(IReadOnlyList<CriticalStudyEntry> studies, Point location, AutomationService? automationService = null, Action<CriticalStudyEntry>? onRemoveStudy = null)
    {
        _studies = studies;
        _automationService = automationService;
        _onRemoveStudy = onRemoveStudy;

        // Form setup
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(35, 35, 35);
        StartPosition = FormStartPosition.Manual;
        Location = location;

        // Calculate size based on number of entries
        var itemHeight = 50;
        var headerHeight = 30;
        var footerHeight = studies.Count > 0 && automationService != null ? 22 : 0;
        var minHeight = 80;
        var maxHeight = 350;
        var width = 280;
        var calculatedHeight = headerHeight + Math.Max(1, studies.Count) * itemHeight + footerHeight + 10;
        var height = Math.Max(minHeight, Math.Min(maxHeight, calculatedHeight));
        Size = new Size(width, height);

        // Border panel
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

        // Header
        var headerPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = headerHeight,
            BackColor = Color.FromArgb(45, 45, 45)
        };

        var headerLabel = new Label
        {
            Text = $"Critical Notes ({studies.Count})",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(255, 120, 120), // Light red
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(10, 0, 0, 0)
        };
        headerPanel.Controls.Add(headerLabel);

        // Close button
        var closeBtn = new Label
        {
            Text = "X",
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = Color.Gray,
            AutoSize = false,
            Size = new Size(24, headerHeight),
            Dock = DockStyle.Right,
            TextAlign = ContentAlignment.MiddleCenter,
            Cursor = Cursors.Hand
        };
        closeBtn.Click += (s, e) => Close();
        closeBtn.MouseEnter += (s, e) => closeBtn.ForeColor = Color.White;
        closeBtn.MouseLeave += (s, e) => closeBtn.ForeColor = Color.Gray;
        headerPanel.Controls.Add(closeBtn);

        // List box
        _listBox = new ListBox
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(40, 40, 40),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.None,
            DrawMode = DrawMode.OwnerDrawFixed,
            ItemHeight = itemHeight,
            IntegralHeight = false
        };
        _listBox.DrawItem += ListBox_DrawItem;
        _listBox.MouseMove += ListBox_MouseMove;
        _listBox.MouseLeave += (s, e) => { _hoveredIndex = -1; _hoveringDelete = false; _listBox.Invalidate(); };
        _listBox.MouseClick += ListBox_MouseClick;
        _listBox.DoubleClick += ListBox_DoubleClick;
        _listBox.Cursor = Cursors.Hand;

        // Populate list
        if (studies.Count == 0)
        {
            _listBox.Items.Add("(no critical notes this session)");
        }
        else
        {
            foreach (var study in studies.OrderByDescending(s => s.CriticalNoteTime))
            {
                _listBox.Items.Add(study);
            }
        }

        // Add controls in correct dock order (WinForms docks in REVERSE order):
        // Fill first (processed last), then Bottom, then Top
        innerPanel.Controls.Add(_listBox);       // Fill - processed last

        // Footer hint (only if we have studies and automation service)
        if (studies.Count > 0 && automationService != null)
        {
            var footerPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 22,
                BackColor = Color.FromArgb(45, 45, 45)
            };

            var footerLabel = new Label
            {
                Text = "Double-click to open in Clario",
                Font = new Font("Segoe UI", 8, FontStyle.Italic),
                ForeColor = Color.FromArgb(120, 120, 120),
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter
            };
            footerPanel.Controls.Add(footerLabel);
            innerPanel.Controls.Add(footerPanel);  // Bottom - processed second
        }

        innerPanel.Controls.Add(headerPanel);      // Top - processed first

        // Handle click outside
        Deactivate += (s, e) =>
        {
            if (_allowDeactivateClose)
                Close();
        };

        // Enable close on deactivate after short delay
        Shown += (s, e) =>
        {
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

    /// <summary>
    /// Returns the delete-button rectangle for a given item bounds.
    /// </summary>
    private Rectangle GetDeleteButtonRect(Rectangle itemBounds)
    {
        return new Rectangle(
            itemBounds.Right - DeleteBtnSize - DeleteBtnMargin,
            itemBounds.Y + (itemBounds.Height - DeleteBtnSize) / 2,
            DeleteBtnSize,
            DeleteBtnSize);
    }

    private void ListBox_DrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= _listBox.Items.Count) return;

        e.DrawBackground();

        var isHovered = e.Index == _hoveredIndex;
        var bgColor = isHovered ? Color.FromArgb(55, 55, 55) : Color.FromArgb(40, 40, 40);

        using var bgBrush = new SolidBrush(bgColor);
        e.Graphics.FillRectangle(bgBrush, e.Bounds);

        var item = _listBox.Items[e.Index];

        if (item is CriticalStudyEntry entry)
        {
            // Line 1: Patient name (bold) + Site code
            var nameText = entry.PatientName;
            var siteText = entry.SiteCode;

            using var nameBrush = new SolidBrush(Color.White);
            using var nameFont = new Font("Segoe UI", 9, FontStyle.Bold);
            e.Graphics.DrawString(nameText, nameFont, nameBrush, e.Bounds.X + 10, e.Bounds.Y + 5);

            // Site code on right side of line 1 (shift left if delete button visible)
            var rightMargin = _onRemoveStudy != null ? DeleteBtnSize + DeleteBtnMargin + 4 : 10;
            using var siteBrush = new SolidBrush(Color.FromArgb(100, 180, 255)); // Light blue
            using var siteFont = new Font("Segoe UI", 8);
            var siteSize = e.Graphics.MeasureString(siteText, siteFont);
            e.Graphics.DrawString(siteText, siteFont, siteBrush, e.Bounds.Right - siteSize.Width - rightMargin, e.Bounds.Y + 6);

            // Line 2: Description + Time
            var descText = TruncateDescription(entry.Description, 25);
            var timeText = entry.CriticalNoteTime.ToString("h:mm tt");

            using var descBrush = new SolidBrush(Color.FromArgb(180, 180, 180));
            using var descFont = new Font("Segoe UI", 8);
            e.Graphics.DrawString(descText, descFont, descBrush, e.Bounds.X + 10, e.Bounds.Y + 26);

            // Time on right side of line 2
            using var timeBrush = new SolidBrush(Color.FromArgb(150, 150, 150));
            var timeSize = e.Graphics.MeasureString(timeText, descFont);
            e.Graphics.DrawString(timeText, descFont, timeBrush, e.Bounds.Right - timeSize.Width - rightMargin, e.Bounds.Y + 26);

            // Delete (X) button — shown on hovered rows
            if (_onRemoveStudy != null && isHovered)
            {
                var btnRect = GetDeleteButtonRect(e.Bounds);
                var btnColor = _hoveringDelete
                    ? Color.FromArgb(255, 80, 80)   // Red when hovering the X
                    : Color.FromArgb(120, 120, 120); // Gray otherwise
                using var xFont = new Font("Segoe UI", 9, FontStyle.Bold);
                using var xBrush = new SolidBrush(btnColor);
                var xSize = e.Graphics.MeasureString("X", xFont);
                e.Graphics.DrawString("X", xFont, xBrush,
                    btnRect.X + (btnRect.Width - xSize.Width) / 2,
                    btnRect.Y + (btnRect.Height - xSize.Height) / 2);
            }

            // Separator line
            if (e.Index < _listBox.Items.Count - 1)
            {
                using var pen = new Pen(Color.FromArgb(60, 60, 60));
                e.Graphics.DrawLine(pen, e.Bounds.X + 10, e.Bounds.Bottom - 1, e.Bounds.Right - 10, e.Bounds.Bottom - 1);
            }
        }
        else
        {
            // Empty message
            using var emptyBrush = new SolidBrush(Color.FromArgb(120, 120, 120));
            using var emptyFont = new Font("Segoe UI", 9, FontStyle.Italic);
            e.Graphics.DrawString(item.ToString() ?? "", emptyFont, emptyBrush, e.Bounds.X + 10, e.Bounds.Y + 15);
        }
    }

    private void ListBox_MouseMove(object? sender, MouseEventArgs e)
    {
        var index = _listBox.IndexFromPoint(e.Location);
        var wasHoveringDelete = _hoveringDelete;

        // Check if mouse is over the delete button area
        if (_onRemoveStudy != null && index >= 0 && index < _listBox.Items.Count && _listBox.Items[index] is CriticalStudyEntry)
        {
            var itemRect = _listBox.GetItemRectangle(index);
            var btnRect = GetDeleteButtonRect(itemRect);
            _hoveringDelete = btnRect.Contains(e.Location);
        }
        else
        {
            _hoveringDelete = false;
        }

        if (index != _hoveredIndex || wasHoveringDelete != _hoveringDelete)
        {
            _hoveredIndex = index;
            _listBox.Invalidate();
            _listBox.Cursor = _hoveringDelete ? Cursors.Hand : Cursors.Default;
        }
    }

    private void ListBox_MouseClick(object? sender, MouseEventArgs e)
    {
        if (_onRemoveStudy == null || e.Button != MouseButtons.Left) return;

        var index = _listBox.IndexFromPoint(e.Location);
        if (index < 0 || index >= _listBox.Items.Count) return;
        if (_listBox.Items[index] is not CriticalStudyEntry) return;

        var itemRect = _listBox.GetItemRectangle(index);
        var btnRect = GetDeleteButtonRect(itemRect);
        if (btnRect.Contains(e.Location))
        {
            _listBox.SelectedIndex = index;
            RemoveSelectedEntry();
        }
    }

    private void ListBox_DoubleClick(object? sender, EventArgs e)
    {
        if (_automationService == null) return;
        if (_listBox.SelectedIndex < 0) return;

        var item = _listBox.Items[_listBox.SelectedIndex];
        if (item is CriticalStudyEntry entry && !string.IsNullOrEmpty(entry.Accession))
        {
            if (string.IsNullOrEmpty(entry.Mrn))
            {
                // Can't open without MRN
                return;
            }

            var success = _automationService.OpenStudyInClario(entry.Accession, entry.Mrn);
            if (success)
            {
                Close();
            }
        }
    }

    private static string TruncateDescription(string desc, int maxLen)
    {
        if (string.IsNullOrEmpty(desc)) return "Unknown";
        if (desc.Length <= maxLen) return desc;
        return desc.Substring(0, maxLen - 3) + "...";
    }

    private void RemoveSelectedEntry()
    {
        if (_listBox.SelectedIndex < 0) return;
        var item = _listBox.Items[_listBox.SelectedIndex];
        if (item is CriticalStudyEntry entry)
        {
            _listBox.Items.RemoveAt(_listBox.SelectedIndex);
            _onRemoveStudy?.Invoke(entry);

            if (_listBox.Items.Count == 0)
            {
                Close();
            }
        }
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == Keys.Escape)
        {
            Close();
            return true;
        }
        if (keyData == Keys.Delete && _onRemoveStudy != null)
        {
            RemoveSelectedEntry();
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }
}
