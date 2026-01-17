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
    public static string Show(string prompt, string title, string defaultValue = "")
    {
        using var form = new Form();
        using var label = new Label();
        using var textBox = new TextBox();
        using var buttonOk = new Button();
        using var buttonCancel = new Button();

        form.Text = title;
        label.Text = prompt;
        textBox.Text = defaultValue;

        buttonOk.Text = "OK";
        buttonCancel.Text = "Cancel";
        buttonOk.DialogResult = DialogResult.OK;
        buttonCancel.DialogResult = DialogResult.Cancel;

        label.SetBounds(12, 12, 372, 20);
        textBox.SetBounds(12, 36, 372, 20);
        buttonOk.SetBounds(228, 72, 75, 23);
        buttonCancel.SetBounds(309, 72, 75, 23);

        label.AutoSize = true;
        textBox.Anchor |= AnchorStyles.Right;
        buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

        form.ClientSize = new Size(396, 107);
        form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
        form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
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
