using System.Windows;
using CommonLibrary.Models;

namespace AuthoringTool.Dialogs;

public partial class EditTaskDialog : Window
{
    private TaskModel _task = new();

    public EditTaskDialog() => InitializeComponent();

    public void SetTask(TaskModel task)
    {
        _task = task;
        TxtObjective.Text = task.Objective;
        TxtInstruction.Text = task.Instruction;
        TxtTemplateFile.Text = task.TemplateFile;
        TxtPoints.Text = task.Points.ToString();
    }

    public TaskModel GetResult()
    {
        var task = new TaskModel
        {
            TaskId = _task.TaskId,
            Number = _task.Number,
            Objective = TxtObjective.Text.Trim(),
            Instruction = TxtInstruction.Text.Trim(),
            TemplateFile = TxtTemplateFile.Text.Trim(),
            Points = int.TryParse(TxtPoints.Text, out var p) ? p : 0,
            ScoringRules = _task.ScoringRules
        };
        return task;
    }

    public void ApplyResult(TaskModel target)
    {
        target.Objective = TxtObjective.Text.Trim();
        target.Instruction = TxtInstruction.Text.Trim();
        target.TemplateFile = TxtTemplateFile.Text.Trim();
        target.Points = int.TryParse(TxtPoints.Text, out var p) ? p : 0;
    }

    private void Browse_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Office Files|*.xlsx;*.xlsm;*.docx;*.pptx|All Files|*.*",
            Title = "Chọn file template cho task"
        };
        if (dlg.ShowDialog() == true)
            TxtTemplateFile.Text = dlg.FileName;
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(TxtInstruction.Text))
        {
            MessageBox.Show("Vui lòng nhập hướng dẫn cho thí sinh.", "Thiếu thông tin",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            TxtInstruction.Focus();
            return;
        }
        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => DialogResult = false;
}
