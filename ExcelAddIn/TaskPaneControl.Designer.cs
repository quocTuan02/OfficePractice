using System.Drawing;
using System.Windows.Forms;

namespace ExcelAddIn
{
    partial class TaskPaneControl
    {
        private System.ComponentModel.IContainer components = null;
        private Label lblTitle;
        private ListBox lstEvents;
        private Button btnClear;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            lblTitle = new Label();
            lstEvents = new ListBox();
            btnClear = new Button();
            SuspendLayout();

            lblTitle.Dock = DockStyle.Top;
            lblTitle.Text = "OfficePractice";
            lblTitle.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
            lblTitle.Height = 32;
            lblTitle.TextAlign = ContentAlignment.MiddleLeft;
            lblTitle.Padding = new Padding(6, 0, 0, 0);
            lblTitle.BackColor = Color.FromArgb(33, 115, 70);
            lblTitle.ForeColor = Color.White;

            lstEvents.Dock = DockStyle.Fill;
            lstEvents.Font = new Font("Consolas", 9f);
            lstEvents.BorderStyle = BorderStyle.None;
            lstEvents.BackColor = Color.FromArgb(30, 30, 46);
            lstEvents.ForeColor = Color.FromArgb(205, 214, 244);
            lstEvents.SelectionMode = SelectionMode.None;

            btnClear.Dock = DockStyle.Bottom;
            btnClear.Text = "Clear Log";
            btnClear.Height = 28;
            btnClear.FlatStyle = FlatStyle.Flat;
            btnClear.BackColor = Color.FromArgb(69, 71, 90);
            btnClear.ForeColor = Color.White;
            btnClear.Click += (s, e) => lstEvents.Items.Clear();

            Controls.Add(lstEvents);
            Controls.Add(lblTitle);
            Controls.Add(btnClear);

            Name = "TaskPaneControl";
            BackColor = Color.FromArgb(24, 24, 37);
            ResumeLayout(false);
        }
    }
}
