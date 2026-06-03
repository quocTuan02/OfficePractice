namespace AuthoringTool.Services;

public static class StorageService
{
    public static string BaseDir { get; }
    public static string TemplatesDir { get; }
    public static string TestsDir { get; }
    public static string PackagesDir { get; }
    public static string TemplateFilesDir { get; }

    static StorageService()
    {
        BaseDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "AuthoringTool");
        TemplatesDir = Path.Combine(BaseDir, "Templates");
        TestsDir = Path.Combine(BaseDir, "Tests");
        PackagesDir = Path.Combine(BaseDir, "Packages");
        TemplateFilesDir = Path.Combine(BaseDir, "TemplateFiles");

        Directory.CreateDirectory(TemplatesDir);
        Directory.CreateDirectory(TestsDir);
        Directory.CreateDirectory(PackagesDir);
        Directory.CreateDirectory(Path.Combine(TemplateFilesDir, "Excel"));
        Directory.CreateDirectory(Path.Combine(TemplateFilesDir, "Word"));
        Directory.CreateDirectory(Path.Combine(TemplateFilesDir, "PowerPoint"));
    }

    public static void Initialize() { }
}
