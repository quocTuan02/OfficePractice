using System.Collections.ObjectModel;
using System.Windows;
using AuthoringTool.Helpers;
using AuthoringTool.Services;
using CommonLibrary.Models;

namespace AuthoringTool.ViewModels;

public class TestPackagerViewModel : ViewModelBase
{
    private readonly TestService _testService = new();
    private readonly PackageService _packageService = new();
    private TestModel? _selectedTest;
    private string _statusMessage = string.Empty;

    public ObservableCollection<TestModel> Tests { get; } = new();
    public ObservableCollection<ValidationIssue> ValidationResults { get; } = new();

    public TestModel? SelectedTest
    {
        get => _selectedTest;
        set
        {
            SetProperty(ref _selectedTest, value);
            OnPropertyChanged(nameof(HasSelection));
            ValidationResults.Clear();
        }
    }

    public bool HasSelection => SelectedTest != null;

    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    public RelayCommand ValidateCommand { get; }
    public RelayCommand ExportZipCommand { get; }
    public RelayCommand ExportFolderCommand { get; }
    public RelayCommand ImportCommand { get; }
    public RelayCommand RefreshCommand { get; }

    public TestPackagerViewModel()
    {
        ValidateCommand = new RelayCommand(Validate, () => HasSelection);
        ExportZipCommand = new RelayCommand(ExportZip, () => HasSelection);
        ExportFolderCommand = new RelayCommand(ExportFolder, () => HasSelection);
        ImportCommand = new RelayCommand(Import);
        RefreshCommand = new RelayCommand(LoadTests);

        LoadTests();
    }

    private void LoadTests()
    {
        Tests.Clear();
        foreach (var t in _testService.LoadAll()) Tests.Add(t);
        StatusMessage = $"Đã tải {Tests.Count} bài thi.";
    }

    private void Validate()
    {
        if (SelectedTest == null) return;
        ValidationResults.Clear();
        var issues = _testService.Validate(SelectedTest, new TemplateService());
        foreach (var issue in issues) ValidationResults.Add(issue);

        if (issues.Count == 0)
        {
            ValidationResults.Add(new ValidationIssue(Severity.Info, "Bài thi hợp lệ, sẵn sàng đóng gói!"));
            StatusMessage = "Validation: OK";
        }
        else
        {
            var errorCount = issues.Count(i => i.Severity == Severity.Error);
            var warnCount = issues.Count(i => i.Severity == Severity.Warning);
            StatusMessage = $"Validation: {errorCount} lỗi, {warnCount} cảnh báo";
        }
    }

    private void ExportZip()
    {
        if (SelectedTest == null) return;
        Validate();
        if (ValidationResults.Any(r => r.Severity == Severity.Error))
        {
            MessageBox.Show("Có lỗi validation. Hãy sửa trước khi đóng gói.",
                "Không thể đóng gói", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            var zipPath = _packageService.ExportZip(SelectedTest);
            StatusMessage = $"Đã xuất: {zipPath}";
            var r = MessageBox.Show($"Đã tạo package:\n{zipPath}\n\nMở thư mục chứa file?",
                "Xuất thành công", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (r == MessageBoxResult.Yes)
                System.Diagnostics.Process.Start("explorer.exe",
                    $"/select,\"{zipPath}\"");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi khi xuất: {ex.Message}", "Lỗi",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ExportFolder()
    {
        if (SelectedTest == null) return;
        Validate();
        if (ValidationResults.Any(r => r.Severity == Severity.Error))
        {
            MessageBox.Show("Có lỗi validation. Hãy sửa trước khi đóng gói.",
                "Không thể đóng gói", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var dialog = new System.Windows.Forms.FolderBrowserDialog
        {
            Description = "Chọn thư mục xuất package"
        };

        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

        try
        {
            var packageDir = _packageService.CreatePackage(SelectedTest, dialog.SelectedPath);
            StatusMessage = $"Đã xuất folder: {packageDir}";
            System.Diagnostics.Process.Start("explorer.exe", $"\"{packageDir}\"");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Import()
    {
        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Package ZIP|*.zip",
            Title = "Chọn file package"
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            var test = _packageService.ImportZip(dlg.FileName);
            _testService.Save(test);
            Tests.Insert(0, test);
            SelectedTest = test;
            StatusMessage = $"Đã import: {test.Name}";
            MessageBox.Show($"Import thành công bài thi: {test.Name}", "Thành công",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lỗi import: {ex.Message}", "Lỗi",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
