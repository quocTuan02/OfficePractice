using System.IO.Compression;
using System.Text.Json;
using CommonLibrary.Models;

namespace AuthoringTool.Services;

public class PackageService
{
    private static readonly JsonSerializerOptions JsonOpts = new() { WriteIndented = true };

    public string CreatePackage(TestModel test, string outputDir)
    {
        var packageName = SanitizeFileName(test.Name) + "_v" + test.Version;
        var packageDir = Path.Combine(outputDir, packageName);
        Directory.CreateDirectory(packageDir);
        Directory.CreateDirectory(Path.Combine(packageDir, "Templates"));
        Directory.CreateDirectory(Path.Combine(packageDir, "ScoringRules"));

        // Write test.json
        File.WriteAllText(
            Path.Combine(packageDir, "test.json"),
            JsonSerializer.Serialize(test, JsonOpts));

        // Copy template files
        foreach (var task in test.Tasks)
        {
            if (!string.IsNullOrEmpty(task.TemplateFile) && File.Exists(task.TemplateFile))
            {
                var dest = Path.Combine(packageDir, "Templates", Path.GetFileName(task.TemplateFile));
                File.Copy(task.TemplateFile, dest, overwrite: true);
            }
        }

        // Write scoring rules per task
        foreach (var task in test.Tasks)
        {
            if (task.ScoringRules.Count > 0)
            {
                File.WriteAllText(
                    Path.Combine(packageDir, "ScoringRules", $"{task.TaskId}.json"),
                    JsonSerializer.Serialize(task.ScoringRules, JsonOpts));
            }
        }

        return packageDir;
    }

    public string ExportZip(TestModel test)
    {
        var tmpDir = Path.Combine(Path.GetTempPath(), "AuthoringTool_Export");
        Directory.CreateDirectory(tmpDir);

        var packageDir = CreatePackage(test, tmpDir);
        var zipPath = Path.Combine(StorageService.PackagesDir,
            SanitizeFileName(test.Name) + "_v" + test.Version + ".zip");

        if (File.Exists(zipPath)) File.Delete(zipPath);
        ZipFile.CreateFromDirectory(packageDir, zipPath);
        Directory.Delete(tmpDir, recursive: true);

        return zipPath;
    }

    public TestModel ImportZip(string zipPath)
    {
        var extractDir = Path.Combine(
            Path.GetTempPath(),
            "AuthoringTool_Import_" + Guid.NewGuid().ToString("N"));
        ZipFile.ExtractToDirectory(zipPath, extractDir);

        var testJsonPath = Directory.GetFiles(extractDir, "test.json", SearchOption.AllDirectories)
            .FirstOrDefault()
            ?? throw new FileNotFoundException("Không tìm thấy test.json trong package.");

        var test = JsonSerializer.Deserialize<TestModel>(File.ReadAllText(testJsonPath))
            ?? throw new InvalidDataException("File test.json không hợp lệ.");

        // Copy template files to local storage
        var templatesDir = Path.GetDirectoryName(testJsonPath)!;
        var localTemplatesDir = Path.Combine(templatesDir, "Templates");
        if (Directory.Exists(localTemplatesDir))
        {
            foreach (var file in Directory.GetFiles(localTemplatesDir))
            {
                var officeType = GuessOfficeType(file);
                var destDir = Path.Combine(StorageService.TemplateFilesDir, officeType);
                Directory.CreateDirectory(destDir);
                File.Copy(file, Path.Combine(destDir, Path.GetFileName(file)), overwrite: true);
            }
        }

        Directory.Delete(extractDir, recursive: true);
        return test;
    }

    private static string GuessOfficeType(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".xlsx" or ".xls" or ".xlsm" => "Excel",
            ".docx" or ".doc" => "Word",
            ".pptx" or ".ppt" => "PowerPoint",
            _ => "Excel"
        };
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        return string.Concat(name.Select(c => invalid.Contains(c) ? '_' : c));
    }
}
