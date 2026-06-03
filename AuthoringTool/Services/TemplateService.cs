using System.Text.Json;
using CommonLibrary.Models;

namespace AuthoringTool.Services;

public class TemplateService
{
    private static readonly JsonSerializerOptions JsonOpts = new() { WriteIndented = true };

    public List<TemplateInfo> LoadAll()
    {
        var list = new List<TemplateInfo>();
        foreach (var file in Directory.GetFiles(StorageService.TemplatesDir, "*.json"))
        {
            try
            {
                var t = JsonSerializer.Deserialize<TemplateInfo>(File.ReadAllText(file));
                if (t != null) list.Add(t);
            }
            catch { /* skip corrupt */ }
        }
        return list.OrderByDescending(t => t.CreatedDate).ToList();
    }

    public void Save(TemplateInfo template)
    {
        var path = Path.Combine(StorageService.TemplatesDir, $"{template.Id}.json");
        File.WriteAllText(path, JsonSerializer.Serialize(template, JsonOpts));
    }

    public void Delete(TemplateInfo template)
    {
        var metaPath = Path.Combine(StorageService.TemplatesDir, $"{template.Id}.json");
        if (File.Exists(metaPath)) File.Delete(metaPath);
    }

    public string GetDefaultFilePath(TemplateInfo template)
    {
        var ext = template.OfficeType switch
        {
            OfficeType.Excel => ".xlsx",
            OfficeType.Word => ".docx",
            OfficeType.PowerPoint => ".pptx",
            _ => ".xlsx"
        };
        return Path.Combine(
            StorageService.TemplateFilesDir,
            template.OfficeType.ToString(),
            template.Id + ext);
    }

    public void OpenWithOffice(TemplateInfo template)
    {
        if (string.IsNullOrEmpty(template.FilePath) || !File.Exists(template.FilePath))
            throw new FileNotFoundException($"File template không tồn tại: {template.FilePath}");

        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        {
            FileName = template.FilePath,
            UseShellExecute = true
        });
    }
}
