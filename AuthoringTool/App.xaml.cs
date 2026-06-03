using System.Windows;
using AuthoringTool.Services;

namespace AuthoringTool;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        StorageService.Initialize();
    }
}
