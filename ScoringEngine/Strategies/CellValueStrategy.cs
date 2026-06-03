using CommonLibrary.Models;
using OfficeOpenXml;
using ScoringEngine.Models;

namespace ScoringEngine.Strategies;

public class CellValueStrategy : IRuleStrategy
{
    public RuleResult Execute(ScoringRule rule, ExcelWorksheet worksheet)
    {
        var result = new RuleResult
        {
            RuleId = rule.RuleId,
            RuleType = nameof(RuleType.CellValue),
            Description = rule.Description,
            PointsMax = rule.Points
        };

        var cellAddr = rule.Parameters.GetValueOrDefault("CellAddress", "A1");
        var expected = rule.Parameters.GetValueOrDefault("ExpectedValue", "");
        var compareType = rule.Parameters.GetValueOrDefault("CompareType", "Equals");
        result.ExpectedValue = expected;

        try
        {
            var cell = worksheet.Cells[cellAddr];
            var actual = cell.Text ?? cell.Value?.ToString() ?? "";
            result.ActualValue = actual;

            result.Passed = compareType switch
            {
                "Equals" => actual.Equals(expected, StringComparison.OrdinalIgnoreCase),
                "Contains" => actual.Contains(expected, StringComparison.OrdinalIgnoreCase),
                "GreaterThan" => double.TryParse(actual, out var a) && double.TryParse(expected, out var e) && a > e,
                "LessThan" => double.TryParse(actual, out var a2) && double.TryParse(expected, out var e2) && a2 < e2,
                "NotEmpty" => !string.IsNullOrWhiteSpace(actual),
                _ => actual.Equals(expected, StringComparison.OrdinalIgnoreCase)
            };

            result.PointsEarned = result.Passed ? rule.Points : 0;
            result.Details = result.Passed
                ? $"Ô {cellAddr} = '{actual}'"
                : $"Ô {cellAddr}: mong đợi '{expected}', thực tế '{actual}'";
        }
        catch (Exception ex)
        {
            result.Details = $"Lỗi đọc ô {cellAddr}: {ex.Message}";
        }

        return result;
    }
}
