using CommonLibrary.Models;
using OfficeOpenXml;
using ScoringEngine.Models;
using ScoringEngine.Strategies;

namespace ScoringEngine;

/// <summary>
/// Entry point for executing scoring rules against an Excel file.
/// Uses Strategy Pattern – each RuleType maps to a dedicated IRuleStrategy.
/// </summary>
public class ScoringExecutor
{
    private readonly Dictionary<RuleType, IRuleStrategy> _strategies;

    public ScoringExecutor()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        _strategies = new Dictionary<RuleType, IRuleStrategy>
        {
            [RuleType.CellValue]        = new CellValueStrategy(),
            [RuleType.Formula]          = new FormulaStrategy(),
            [RuleType.Formatting]       = new FormattingStrategy(),
            [RuleType.ObjectExistence]  = new ObjectExistenceStrategy(),
            [RuleType.RangeComparison]  = new RangeComparisonStrategy(),
        };
    }

    // ── Single-rule execution (used by "Test Rule" button in AuthoringTool) ──

    public RuleResult ExecuteRule(ScoringRule rule, string filePath, int sheetIndex = 0)
    {
        using var pkg = new ExcelPackage(new FileInfo(filePath));
        var ws = GetSheet(pkg, sheetIndex);
        return _strategies[rule.RuleType].Execute(rule, ws);
    }

    // ── Full-task execution (used by MainApp on submission) ──

    public TaskScoreResult ExecuteTask(TaskModel task, string filePath, int sheetIndex = 0)
    {
        using var pkg = new ExcelPackage(new FileInfo(filePath));
        var ws = GetSheet(pkg, sheetIndex);

        var result = new TaskScoreResult
        {
            TaskId        = task.TaskId,
            TaskNumber    = task.Number,
            TaskObjective = task.Objective
        };

        foreach (var rule in task.ScoringRules)
            result.RuleResults.Add(_strategies[rule.RuleType].Execute(rule, ws));

        return result;
    }

    // ── Whole-test execution ──

    public List<TaskScoreResult> ExecuteTest(TestModel test, string workDir)
    {
        return test.Tasks
            .Select(task =>
            {
                var filePath = string.IsNullOrEmpty(task.TemplateFile)
                    ? null
                    : Path.IsPathRooted(task.TemplateFile)
                        ? task.TemplateFile
                        : Path.Combine(workDir, task.TemplateFile);

                if (filePath == null || !File.Exists(filePath))
                    return new TaskScoreResult
                    {
                        TaskId        = task.TaskId,
                        TaskNumber    = task.Number,
                        TaskObjective = task.Objective,
                        RuleResults   = task.ScoringRules.Select(r => new RuleResult
                        {
                            RuleId      = r.RuleId,
                            Description = r.Description,
                            PointsMax   = r.Points,
                            Details     = $"File không tồn tại: {filePath}"
                        }).ToList()
                    };

                return ExecuteTask(task, filePath);
            })
            .ToList();
    }

    private static ExcelWorksheet GetSheet(ExcelPackage pkg, int index)
        => index < pkg.Workbook.Worksheets.Count
            ? pkg.Workbook.Worksheets[index]
            : pkg.Workbook.Worksheets.First();
}
