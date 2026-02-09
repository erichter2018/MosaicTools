using System;
using System.Drawing;
using System.Windows.Forms;

namespace MosaicTools.Services;

/// <summary>
/// A native WinForms replacement for the legacy Interaction.InputBox.
/// Prevents PlatformNotSupportedException.
/// </summary>
public static class InputBox
{
    public static string Show(string prompt, string title, string defaultValue = "", bool multiline = false)
    {
        using var form = new Form();
        var label = new Label();
        var textBox = new TextBox();
        var buttonOk = new Button();
        var buttonCancel = new Button();

        form.Text = title;
        label.Text = prompt;
        textBox.Text = defaultValue;
        textBox.SelectionStart = 0;
        textBox.SelectionLength = 0;

        buttonOk.Text = "OK";
        buttonCancel.Text = "Cancel";
        buttonOk.DialogResult = DialogResult.OK;
        buttonCancel.DialogResult = DialogResult.Cancel;

        label.AutoSize = true;

        if (multiline)
        {
            textBox.Multiline = true;
            textBox.ScrollBars = ScrollBars.Vertical;
            textBox.AcceptsReturn = true;
            textBox.WordWrap = false;
            label.SetBounds(12, 12, 450, 20);
            textBox.SetBounds(12, 40, 450, 180);
            buttonOk.SetBounds(306, 230, 75, 23);
            buttonCancel.SetBounds(387, 230, 75, 23);
            form.ClientSize = new Size(480, 265);
        }
        else
        {
            label.SetBounds(12, 12, 372, 20);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);
            form.ClientSize = new Size(396, 107);
        }

        textBox.Anchor |= AnchorStyles.Right;
        buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

        form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
        if (!multiline)
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
        if (!multiline)
            form.AcceptButton = buttonOk;
        form.CancelButton = buttonCancel;

        // Visual Polish
        form.BackColor = Color.FromArgb(45, 45, 45);
        form.ForeColor = Color.White;
        textBox.BackColor = Color.FromArgb(60, 60, 60);
        textBox.ForeColor = Color.White;
        buttonOk.FlatStyle = FlatStyle.Flat;
        buttonCancel.FlatStyle = FlatStyle.Flat;
        buttonOk.BackColor = Color.FromArgb(70, 70, 70);
        buttonCancel.BackColor = Color.FromArgb(70, 70, 70);

        if (form.ShowDialog() == DialogResult.OK)
        {
            return textBox.Text;
        }

        return defaultValue;
    }
}
