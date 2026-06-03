using CommonLibrary.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using ScoringEngine.Models;

namespace ScoringEngine.Strategies;

public class FormattingStrategy : IRuleStrategy
{
    public RuleResult Execute(ScoringRule rule, ExcelWorksheet worksheet)
    {
        var result = new RuleResult
        {
            RuleId = rule.RuleId,
            RuleType = nameof(RuleType.Formatting),
            Description = rule.Description,
            PointsMax = rule.Points
        };

        var cellAddr = rule.Parameters.GetValueOrDefault("CellAddress", "A1");
        var property = rule.Parameters.GetValueOrDefault("Property", "Bold");
        var expectedValue = rule.Parameters.GetValueOrDefault("ExpectedValue", "").Trim();
        result.ExpectedValue = expectedValue;

        try
        {
            var cell = worksheet.Cells[cellAddr];
            var actualValue = GetFormattingValue(cell, property);
            result.ActualValue = actualValue;

            result.Passed = NormalizeForCompare(actualValue)
                .Equals(NormalizeForCompare(expectedValue), StringComparison.OrdinalIgnoreCase);
            result.PointsEarned = result.Passed ? rule.Points : 0;
            result.Details = result.Passed
                ? $"{property} của {cellAddr} = '{actualValue}'"
                : $"{property} của {cellAddr}: mong đợi '{expectedValue}', thực tế '{actualValue}'";
        }
        catch (Exception ex)
        {
            result.Details = $"Lỗi đọc định dạng {cellAddr}: {ex.Message}";
        }

        return result;
    }

    private static string GetFormattingValue(ExcelRange cell, string property) => property switch
    {
        "Bold" => cell.Style.Font.Bold.ToString(),
        "Italic" => cell.Style.Font.Italic.ToString(),
        "Underline" => (cell.Style.Font.UnderLineType != ExcelUnderLineType.None).ToString(),
        "FontSize" => ((int)cell.Style.Font.Size).ToString(),
        "FontColor" => NormalizeRgb(cell.Style.Font.Color.Rgb),
        "BackgroundColor" => NormalizeRgb(cell.Style.Fill.BackgroundColor.Rgb),
        "NumberFormat" => cell.Style.Numberformat.Format ?? "General",
        "WrapText" => cell.Style.WrapText.ToString(),
        _ => "Unknown"
    };

    // EPPlus returns AARRGGBB (8 chars), normalize to #RRGGBB
    private static string NormalizeRgb(string? rgb)
    {
        if (string.IsNullOrEmpty(rgb)) return "#000000";
        rgb = rgb.TrimStart('#');
        return "#" + (rgb.Length == 8 ? rgb[2..] : rgb.Length == 6 ? rgb : rgb.PadLeft(6, '0'));
    }

    // Normalize expected value: "#FF0000" → "ff0000", "True" → "true"
    private static string NormalizeForCompare(string value)
    {
        value = value.TrimStart('#').ToLower();
        // Map AARRGGBB → RRGGBB for comparison
        if (value.Length == 8 && value.All(c => Uri.IsHexDigit(c)))
            value = value[2..];
        return value;
    }
}
