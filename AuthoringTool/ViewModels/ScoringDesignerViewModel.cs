using System.Collections.ObjectModel;
using System.Windows;
using AuthoringTool.Dialogs;
using AuthoringTool.Helpers;
using AuthoringTool.Services;
using CommonLibrary.Models;
using ScoringEngine;
using ScoringEngine.Models;

namespace AuthoringTool.ViewModels;

public class ScoringDesignerViewModel : ViewModelBase
{
    private readonly TestService _testService = new();
    private readonly ScoringExecutor _executor = new();

    private TestModel? _selectedTest;
    private TaskModel? _selectedTask;
    private ScoringRule? _selectedRule;
    private string _statusMessage = string.Empty;
    private string _testFilePath = string.Empty;
    private bool _isTestPanelVisible;

    public ObservableCollection<TestModel>  Tests { get; } = new();
    public ObservableCollection<TaskModel>  Tasks { get; } = new();
    public ObservableCollection<ScoringRule> Rules { get; } = new();
    public ObservableCollection<RuleResult> TestResults { get; } = new();

    // ── Selection ──────────────────────────────────────────────────

    public TestModel? SelectedTest
    {
        get => _selectedTest;
        set { SetProperty(ref _selectedTest, value); LoadTasks(); }
    }

    public TaskModel? SelectedTask
    {
        get => _selectedTask;
        set
        {
            SetProperty(ref _selectedTask, value);
            OnPropertyChanged(nameof(HasTaskSelection));
            LoadRules();
        }
    }

    public ScoringRule? SelectedRule
    {
        get => _selectedRule;
        set
        {
            SetProperty(ref _selectedRule, value);
            OnPropertyChanged(nameof(HasRuleSelection));
            OnPropertyChanged(nameof(SelectedRuleParameters));
        }
    }

    public bool HasTaskSelection => SelectedTask != null;
    public bool HasRuleSelection => SelectedRule != null;

    public string SelectedRuleParameters =>
        SelectedRule == null ? string.Empty
        : string.Join("\n", SelectedRule.Parameters.Select(kv => $"  {kv.Key}: {kv.Value}"));

    // ── Test Panel ─────────────────────────────────────────────────

    public string TestFilePath
    {
        get => _testFilePath;
        set { SetProperty(ref _testFilePath, value); OnPropertyChanged(nameof(HasTestFile)); }
    }

    public bool HasTestFile => !string.IsNullOrEmpty(_testFilePath) && System.IO.File.Exists(_testFilePath);

    public bool IsTestPanelVisible
    {
        get => _isTestPanelVisible;
        set => SetProperty(ref _isTestPanelVisible, value);
    }

    public int TestTotalEarned => TestResults.Sum(r => r.PointsEarned);
    public int TestTotalMax    => TestResults.Sum(r => r.PointsMax);
    public string TestSummary  => TestResults.Count == 0
        ? string.Empty
        : $"{TestResults.Count(r => r.Passed)}/{TestResults.Count} rules đúng · {TestTotalEarned}/{TestTotalMax} điểm";

    // ── Status ─────────────────────────────────────────────────────

    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    // ── Commands ───────────────────────────────────────────────────

    public RelayCommand AddRuleCommand        { get; }
    public RelayCommand EditRuleCommand       { get; }
    public RelayCommand DeleteRuleCommand     { get; }
    public RelayCommand BrowseTestFileCommand { get; }
    public RelayCommand TestRuleCommand       { get; }
    public RelayCommand TestAllRulesCommand   { get; }

    public ScoringDesignerViewModel()
    {
        AddRuleCommand        = new RelayCommand(AddRule,        () => HasTaskSelection);
        EditRuleCommand       = new RelayCommand(EditRule,       () => HasRuleSelection);
        DeleteRuleCommand     = new RelayCommand(DeleteRule,     () => HasRuleSelection);
        BrowseTestFileCommand = new RelayCommand(BrowseTestFile, () => HasTaskSelection);
        TestRuleCommand       = new RelayCommand(TestRule,       () => HasRuleSelection && HasTestFile);
        TestAllRulesCommand   = new RelayCommand(TestAllRules,   () => HasTaskSelection && HasTestFile);

        LoadTests();
    }

    // ── Load ───────────────────────────────────────────────────────

    private void LoadTests()
    {
        Tests.Clear();
        foreach (var t in _testService.LoadAll()) Tests.Add(t);
        SelectedTest = Tests.FirstOrDefault();
    }

    private void LoadTasks()
    {
        Tasks.Clear();
        Rules.Clear();
        TestResults.Clear();
        IsTestPanelVisible = false;
        if (SelectedTest == null) return;
        foreach (var t in SelectedTest.Tasks.OrderBy(t => t.Number)) Tasks.Add(t);
        SelectedTask = Tasks.FirstOrDefault();
    }

    private void LoadRules()
    {
        Rules.Clear();
        TestResults.Clear();
        IsTestPanelVisible = false;
        if (SelectedTask == null) return;
        foreach (var r in SelectedTask.ScoringRules) Rules.Add(r);
        SelectedRule = Rules.FirstOrDefault();
    }

    // ── Rule CRUD ──────────────────────────────────────────────────

    private void AddRule()
    {
        if (SelectedTask == null) return;
        var dialog = new EditRuleDialog { Owner = Application.Current.MainWindow };
        dialog.SetRule(new ScoringRule { TaskId = SelectedTask.TaskId });
        if (dialog.ShowDialog() != true) return;

        var rule = dialog.GetResult();
        SelectedTask.ScoringRules.Add(rule);
        _testService.Save(SelectedTest!);
        Rules.Add(rule);
        SelectedRule = rule;
        StatusMessage = $"Đã thêm rule: {rule.Description}";
    }

    private void EditRule()
    {
        if (SelectedRule == null || SelectedTask == null) return;
        var dialog = new EditRuleDialog { Owner = Application.Current.MainWindow };
        dialog.SetRule(SelectedRule);
        if (dialog.ShowDialog() != true) return;

        dialog.ApplyResult(SelectedRule);
        _testService.Save(SelectedTest!);

        var idx = Rules.IndexOf(SelectedRule);
        Rules.RemoveAt(idx);
        Rules.Insert(idx, SelectedRule);
        SelectedRule = Rules[idx];
        OnPropertyChanged(nameof(SelectedRuleParameters));
        StatusMessage = "Đã cập nhật rule.";
    }

    private void DeleteRule()
    {
        if (SelectedRule == null || SelectedTask == null) return;
        if (MessageBox.Show("Xóa rule này?", "Xác nhận",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;

        SelectedTask.ScoringRules.Remove(SelectedRule);
        _testService.Save(SelectedTest!);
        Rules.Remove(SelectedRule);
        StatusMessage = "Đã xóa rule.";
    }

    // ── Test execution ─────────────────────────────────────────────

    private void BrowseTestFile()
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Files|*.xlsx;*.xlsm;*.xls|All Files|*.*",
            Title  = "Chọn file bài làm cần chấm"
        };
        if (dlg.ShowDialog() != true) return;
        TestFilePath = dlg.FileName;
        StatusMessage = $"File chấm: {System.IO.Path.GetFileName(TestFilePath)}";
    }

    private void TestRule()
    {
        if (SelectedRule == null || !HasTestFile) return;
        RunTests([SelectedRule], $"Rule: {SelectedRule.Description}");
    }

    private void TestAllRules()
    {
        if (SelectedTask == null || !HasTestFile) return;
        RunTests(SelectedTask.ScoringRules, $"Task {SelectedTask.Number} – tất cả {SelectedTask.ScoringRules.Count} rules");
    }

    private void RunTests(IEnumerable<ScoringRule> rules, string label)
    {
        TestResults.Clear();
        IsTestPanelVisible = true;

        try
        {
            foreach (var rule in rules)
            {
                var result = _executor.ExecuteRule(rule, TestFilePath);
                TestResults.Add(result);
            }

            OnPropertyChanged(nameof(TestTotalEarned));
            OnPropertyChanged(nameof(TestTotalMax));
            OnPropertyChanged(nameof(TestSummary));
            StatusMessage = $"{label} → {TestSummary}";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Lỗi chấm điểm: {ex.Message}";
            MessageBox.Show($"Không thể đọc file:\n{ex.Message}", "Lỗi",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
