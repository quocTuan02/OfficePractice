using System;
using System.Windows.Forms;

namespace ExcelAddIn
{
    public partial class TaskPaneControl : UserControl
    {
        public TaskPaneControl()
        {
            InitializeComponent();
        }

        public void AddLog(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AddLog), message);
                return;
            }
            var entry = $"[{DateTime.Now:HH:mm:ss}] {message}";
            lstEvents.Items.Insert(0, entry);
            if (lstEvents.Items.Count > 200)
                lstEvents.Items.RemoveAt(200);
        }
    }
}
