using System.Text.Json;
using CommonLibrary.Models;

namespace AuthoringTool.Services;

public class TestService
{
    private static readonly JsonSerializerOptions JsonOpts = new() { WriteIndented = true };

    public List<TestModel> LoadAll()
    {
        var list = new List<TestModel>();
        foreach (var file in Directory.GetFiles(StorageService.TestsDir, "*.json"))
        {
            try
            {
                var t = JsonSerializer.Deserialize<TestModel>(File.ReadAllText(file));
                if (t != null) list.Add(t);
            }
            catch { /* skip corrupt */ }
        }
        return list.OrderByDescending(t => t.CreatedDate).ToList();
    }

    public TestModel? Load(string id)
    {
        var path = Path.Combine(StorageService.TestsDir, $"{id}.json");
        if (!File.Exists(path)) return null;
        return JsonSerializer.Deserialize<TestModel>(File.ReadAllText(path));
    }

    public void Save(TestModel test)
    {
        var path = Path.Combine(StorageService.TestsDir, $"{test.Id}.json");
        File.WriteAllText(path, JsonSerializer.Serialize(test, JsonOpts));
    }

    public void Delete(string id)
    {
        var path = Path.Combine(StorageService.TestsDir, $"{id}.json");
        if (File.Exists(path)) File.Delete(path);
    }

    public List<ValidationIssue> Validate(TestModel test, TemplateService templateService)
    {
        var issues = new List<ValidationIssue>();

        if (string.IsNullOrWhiteSpace(test.Name))
            issues.Add(new ValidationIssue(Severity.Error, "Bài thi chưa có tên."));

        if (test.Tasks.Count == 0)
            issues.Add(new ValidationIssue(Severity.Error, "Bài thi chưa có task nào."));

        foreach (var task in test.Tasks)
        {
            if (string.IsNullOrWhiteSpace(task.Instruction))
                issues.Add(new ValidationIssue(Severity.Warning, $"Task {task.Number}: Chưa có hướng dẫn."));

            if (string.IsNullOrWhiteSpace(task.TemplateFile))
                issues.Add(new ValidationIssue(Severity.Error, $"Task {task.Number}: Chưa gán template file."));

            if (task.ScoringRules.Count == 0)
                issues.Add(new ValidationIssue(Severity.Warning, $"Task {task.Number}: Chưa có scoring rule."));

            if (task.Points <= 0)
                issues.Add(new ValidationIssue(Severity.Warning, $"Task {task.Number}: Điểm = 0."));
        }

        return issues;
    }
}

public enum Severity { Info, Warning, Error }

public class ValidationIssue
{
    public Severity Severity { get; }
    public string Message { get; }
    public string Icon => Severity switch
    {
        Severity.Error => "✗",
        Severity.Warning => "⚠",
        _ => "ℹ"
    };

    public ValidationIssue(Severity severity, string message)
    {
        Severity = severity;
        Message = message;
    }
}
