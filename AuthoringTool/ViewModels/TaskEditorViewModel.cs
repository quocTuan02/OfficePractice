using System.Collections.ObjectModel;
using System.Windows;
using AuthoringTool.Dialogs;
using AuthoringTool.Helpers;
using AuthoringTool.Services;
using CommonLibrary.Models;

namespace AuthoringTool.ViewModels;

public class TaskEditorViewModel : ViewModelBase
{
    private readonly TestService _testService = new();
    private TestModel? _selectedTest;
    private TaskModel? _selectedTask;
    private string _statusMessage = string.Empty;

    public ObservableCollection<TestModel> Tests { get; } = new();
    public ObservableCollection<TaskModel> Tasks { get; } = new();

    public TestModel? SelectedTest
    {
        get => _selectedTest;
        set
        {
            SetProperty(ref _selectedTest, value);
            OnPropertyChanged(nameof(HasTestSelection));
            LoadTasks();
        }
    }

    public TaskModel? SelectedTask
    {
        get => _selectedTask;
        set
        {
            SetProperty(ref _selectedTask, value);
            OnPropertyChanged(nameof(HasTaskSelection));
        }
    }

    public bool HasTestSelection => SelectedTest != null;
    public bool HasTaskSelection => SelectedTask != null;

    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    public RelayCommand NewTestCommand { get; }
    public RelayCommand DeleteTestCommand { get; }
    public RelayCommand AddTaskCommand { get; }
    public RelayCommand EditTaskCommand { get; }
    public RelayCommand DeleteTaskCommand { get; }
    public RelayCommand MoveUpCommand { get; }
    public RelayCommand MoveDownCommand { get; }
    public RelayCommand SaveTestCommand { get; }

    public TaskEditorViewModel()
    {
        NewTestCommand = new RelayCommand(CreateTest);
        DeleteTestCommand = new RelayCommand(DeleteTest, () => HasTestSelection);
        AddTaskCommand = new RelayCommand(AddTask, () => HasTestSelection);
        EditTaskCommand = new RelayCommand(EditTask, () => HasTaskSelection);
        DeleteTaskCommand = new RelayCommand(DeleteTask, () => HasTaskSelection);
        MoveUpCommand = new RelayCommand(MoveUp, () => CanMoveUp());
        MoveDownCommand = new RelayCommand(MoveDown, () => CanMoveDown());
        SaveTestCommand = new RelayCommand(SaveTest, () => HasTestSelection);

        LoadTests();
    }

    private void LoadTests()
    {
        Tests.Clear();
        foreach (var t in _testService.LoadAll())
            Tests.Add(t);
    }

    private void LoadTasks()
    {
        Tasks.Clear();
        if (SelectedTest == null) return;
        foreach (var t in SelectedTest.Tasks.OrderBy(t => t.Number))
            Tasks.Add(t);
        SelectedTask = Tasks.FirstOrDefault();
    }

    private void CreateTest()
    {
        var dialog = new NewTestDialog { Owner = Application.Current.MainWindow };
        if (dialog.ShowDialog() != true) return;

        var test = new TestModel
        {
            Name = dialog.TestName,
            Description = dialog.TestDescription,
            TotalTimeMinutes = dialog.TotalTimeMinutes
        };

        _testService.Save(test);
        Tests.Insert(0, test);
        SelectedTest = test;
        StatusMessage = $"Đã tạo bài thi: {test.Name}";
    }

    private void DeleteTest()
    {
        if (SelectedTest == null) return;
        var r = MessageBox.Show($"Xóa bài thi '{SelectedTest.Name}'?", "Xác nhận",
            MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (r != MessageBoxResult.Yes) return;

        _testService.Delete(SelectedTest.Id);
        Tests.Remove(SelectedTest);
        SelectedTest = Tests.FirstOrDefault();
        StatusMessage = "Đã xóa bài thi.";
    }

    private void AddTask()
    {
        if (SelectedTest == null) return;

        var dialog = new EditTaskDialog { Owner = Application.Current.MainWindow };
        dialog.SetTask(new TaskModel { Number = Tasks.Count + 1 });
        if (dialog.ShowDialog() != true) return;

        var task = dialog.GetResult();
        task.Number = Tasks.Count + 1;
        SelectedTest.Tasks.Add(task);
        _testService.Save(SelectedTest);

        Tasks.Add(task);
        SelectedTask = task;
        StatusMessage = $"Đã thêm Task {task.Number}.";
    }

    private void EditTask()
    {
        if (SelectedTask == null) return;

        var dialog = new EditTaskDialog { Owner = Application.Current.MainWindow };
        dialog.SetTask(SelectedTask);
        if (dialog.ShowDialog() != true) return;

        dialog.ApplyResult(SelectedTask);
        _testService.Save(SelectedTest!);

        var idx = Tasks.IndexOf(SelectedTask);
        Tasks.RemoveAt(idx);
        Tasks.Insert(idx, SelectedTask);
        SelectedTask = Tasks[idx];
        StatusMessage = $"Đã cập nhật Task {SelectedTask.Number}.";
    }

    private void DeleteTask()
    {
        if (SelectedTask == null || SelectedTest == null) return;
        var r = MessageBox.Show($"Xóa Task {SelectedTask.Number}?", "Xác nhận",
            MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (r != MessageBoxResult.Yes) return;

        SelectedTest.Tasks.Remove(SelectedTask);
        Tasks.Remove(SelectedTask);
        RenumberTasks();
        _testService.Save(SelectedTest);
        StatusMessage = "Đã xóa task.";
    }

    private void RenumberTasks()
    {
        for (int i = 0; i < Tasks.Count; i++)
            Tasks[i].Number = i + 1;
    }

    private bool CanMoveUp() => SelectedTask != null && Tasks.IndexOf(SelectedTask) > 0;
    private bool CanMoveDown() => SelectedTask != null && Tasks.IndexOf(SelectedTask) < Tasks.Count - 1;

    private void MoveUp()
    {
        if (SelectedTask == null) return;
        var idx = Tasks.IndexOf(SelectedTask);
        if (idx <= 0) return;
        Tasks.Move(idx, idx - 1);
        SelectedTest!.Tasks = Tasks.ToList();
        RenumberTasks();
        _testService.Save(SelectedTest);
    }

    private void MoveDown()
    {
        if (SelectedTask == null) return;
        var idx = Tasks.IndexOf(SelectedTask);
        if (idx >= Tasks.Count - 1) return;
        Tasks.Move(idx, idx + 1);
        SelectedTest!.Tasks = Tasks.ToList();
        RenumberTasks();
        _testService.Save(SelectedTest);
    }

    private void SaveTest()
    {
        if (SelectedTest == null) return;
        _testService.Save(SelectedTest);
        StatusMessage = "Đã lưu bài thi.";
    }
}
