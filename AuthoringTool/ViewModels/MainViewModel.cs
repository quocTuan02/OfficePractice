using AuthoringTool.Helpers;

namespace AuthoringTool.ViewModels;

public class MainViewModel : ViewModelBase
{
    private ViewModelBase _currentModule;

    public TemplateManagerViewModel TemplateManagerVM { get; } = new();
    public TaskEditorViewModel TaskEditorVM { get; } = new();
    public ScoringDesignerViewModel ScoringDesignerVM { get; } = new();
    public TestPackagerViewModel TestPackagerVM { get; } = new();

    public ViewModelBase CurrentModule
    {
        get => _currentModule;
        private set => SetProperty(ref _currentModule, value);
    }

    public string CurrentModuleName { get; private set; } = "Templates";

    public RelayCommand NavigateTemplatesCommand { get; }
    public RelayCommand NavigateTasksCommand { get; }
    public RelayCommand NavigateScoringCommand { get; }
    public RelayCommand NavigatePackagerCommand { get; }

    public MainViewModel()
    {
        _currentModule = TemplateManagerVM;

        NavigateTemplatesCommand = new RelayCommand(() => Navigate("Templates"));
        NavigateTasksCommand = new RelayCommand(() => Navigate("Tasks"));
        NavigateScoringCommand = new RelayCommand(() => Navigate("Scoring"));
        NavigatePackagerCommand = new RelayCommand(() => Navigate("Packager"));
    }

    private void Navigate(string module)
    {
        CurrentModuleName = module;
        CurrentModule = module switch
        {
            "Templates" => TemplateManagerVM,
            "Tasks" => TaskEditorVM,
            "Scoring" => ScoringDesignerVM,
            "Packager" => TestPackagerVM,
            _ => TemplateManagerVM
        };
        OnPropertyChanged(nameof(CurrentModuleName));
    }
}
