using System;
using System.Drawing;
using System.Windows.Forms;

namespace RedInk.OutlookAddIn.TaskPane
{
    public class AssistantTaskPaneControl : UserControl
    {
        private readonly Label _titleLabel;
        private readonly TextBox _contentBox;
        private readonly Label _statusLabel;

        public AssistantTaskPaneControl()
        {
            Dock = DockStyle.Fill;
            Padding = new Padding(12);

            _titleLabel = new Label
            {
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                Text = "KI-Assistent",
                Height = 32
            };

            _contentBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new Font("Segoe UI", 10),
                BackColor = Color.White,
                ForeColor = Color.Black
            };

            _statusLabel = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.DimGray
            };

            Controls.Add(_contentBox);
            Controls.Add(_statusLabel);
            Controls.Add(_titleLabel);
        }

        public void DisplayResponse(string intentLabel, string response)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => DisplayResponse(intentLabel, response)));
                return;
            }

            _titleLabel.Text = $"KI-Assistent – {intentLabel}";
            _contentBox.Text = response ?? string.Empty;
            _statusLabel.Text = "Fertig";
        }

        public void ShowStatus(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => ShowStatus(message)));
                return;
            }

            _statusLabel.Text = message ?? string.Empty;

            if (!string.IsNullOrEmpty(message))
            {
                _contentBox.Text = string.Empty;
            }
        }
    }
}
