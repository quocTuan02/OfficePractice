using CommonLibrary.Models;
using OfficeOpenXml;
using ScoringEngine.Models;

namespace ScoringEngine.Strategies;

public class FormulaStrategy : IRuleStrategy
{
    public RuleResult Execute(ScoringRule rule, ExcelWorksheet worksheet)
    {
        var result = new RuleResult
        {
            RuleId = rule.RuleId,
            RuleType = nameof(RuleType.Formula),
            Description = rule.Description,
            PointsMax = rule.Points
        };

        var cellAddr = rule.Parameters.GetValueOrDefault("CellAddress", "A1");
        // Normalize: strip leading '=' for comparison
        var expectedFormula = rule.Parameters.GetValueOrDefault("ExpectedFormula", "").TrimStart('=');
        result.ExpectedValue = $"={expectedFormula}";

        try
        {
            var cell = worksheet.Cells[cellAddr];
            var actualFormula = cell.Formula ?? "";
            result.ActualValue = string.IsNullOrEmpty(actualFormula)
                ? $"(giá trị: {cell.Text})"
                : $"={actualFormula}";

            result.Passed = actualFormula.Equals(expectedFormula, StringComparison.OrdinalIgnoreCase);
            result.PointsEarned = result.Passed ? rule.Points : 0;
            result.Details = result.Passed
                ? $"Ô {cellAddr} có công thức đúng: {result.ActualValue}"
                : $"Ô {cellAddr}: mong đợi '={expectedFormula}', thực tế '{result.ActualValue}'";
        }
        catch (Exception ex)
        {
            result.Details = $"Lỗi đọc ô {cellAddr}: {ex.Message}";
        }

        return result;
    }
}
