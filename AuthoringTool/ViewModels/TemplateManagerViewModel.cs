using System.Collections.ObjectModel;
using System.Windows;
using AuthoringTool.Dialogs;
using AuthoringTool.Helpers;
using AuthoringTool.Services;
using CommonLibrary.Models;

namespace AuthoringTool.ViewModels;

public class TemplateManagerViewModel : ViewModelBase
{
    private readonly TemplateService _service = new();
    private TemplateInfo? _selectedTemplate;
    private string _statusMessage = string.Empty;

    public ObservableCollection<TemplateInfo> Templates { get; } = new();

    public TemplateInfo? SelectedTemplate
    {
        get => _selectedTemplate;
        set
        {
            SetProperty(ref _selectedTemplate, value);
            OnPropertyChanged(nameof(HasSelection));
        }
    }

    public bool HasSelection => SelectedTemplate != null;

    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    public RelayCommand NewCommand { get; }
    public RelayCommand DeleteCommand { get; }
    public RelayCommand OpenInOfficeCommand { get; }
    public RelayCommand BrowseFileCommand { get; }
    public RelayCommand SaveMetaCommand { get; }
    public RelayCommand RefreshCommand { get; }

    public TemplateManagerViewModel()
    {
        NewCommand = new RelayCommand(CreateNew);
        DeleteCommand = new RelayCommand(Delete, () => HasSelection);
        OpenInOfficeCommand = new RelayCommand(OpenInOffice, () => HasSelection);
        BrowseFileCommand = new RelayCommand(BrowseFile, () => HasSelection);
        SaveMetaCommand = new RelayCommand(SaveMeta, () => HasSelection);
        RefreshCommand = new RelayCommand(LoadTemplates);

        LoadTemplates();
    }

    private void LoadTemplates()
    {
        Templates.Clear();
        foreach (var t in _service.LoadAll())
            Templates.Add(t);
        StatusMessage = $"Đã tải {Templates.Count} template.";
    }

    private void CreateNew()
    {
        var dialog = new NewTemplateDialog { Owner = Application.Current.MainWindow };
        if (dialog.ShowDialog() != true) return;

        var template = new TemplateInfo
        {
            Name = dialog.TemplateName,
            Description = dialog.TemplateDescription,
            OfficeType = dialog.SelectedOfficeType
        };

        _service.Save(template);
        Templates.Insert(0, template);
        SelectedTemplate = template;
        StatusMessage = $"Đã tạo template: {template.Name}";
    }

    private void Delete()
    {
        if (SelectedTemplate == null) return;
        var result = MessageBox.Show(
            $"Xóa template '{SelectedTemplate.Name}'?",
            "Xác nhận xóa",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes) return;

        _service.Delete(SelectedTemplate);
        Templates.Remove(SelectedTemplate);
        SelectedTemplate = null;
        StatusMessage = "Đã xóa template.";
    }

    private void OpenInOffice()
    {
        if (SelectedTemplate == null) return;
        try
        {
            _service.OpenWithOffice(SelectedTemplate);
            StatusMessage = $"Đã mở: {SelectedTemplate.FilePath}";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BrowseFile()
    {
        if (SelectedTemplate == null) return;

        var filter = SelectedTemplate.OfficeType switch
        {
            OfficeType.Excel => "Excel Files|*.xlsx;*.xlsm;*.xls",
            OfficeType.Word => "Word Files|*.docx;*.doc",
            OfficeType.PowerPoint => "PowerPoint Files|*.pptx;*.ppt",
            _ => "All Files|*.*"
        };

        var dlg = new Microsoft.Win32.OpenFileDialog
        {
            Filter = filter + "|All Files|*.*",
            Title = "Chọn file template"
        };

        if (dlg.ShowDialog() != true) return;

        SelectedTemplate.FilePath = dlg.FileName;
        _service.Save(SelectedTemplate);
        OnPropertyChanged(nameof(SelectedTemplate));
        StatusMessage = $"Đã gán file: {dlg.FileName}";
    }

    private void SaveMeta()
    {
        if (SelectedTemplate == null) return;
        _service.Save(SelectedTemplate);
        StatusMessage = "Đã lưu thông tin template.";
    }
}
