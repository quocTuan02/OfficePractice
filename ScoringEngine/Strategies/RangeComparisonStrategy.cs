using CommonLibrary.Models;
using OfficeOpenXml;
using ScoringEngine.Models;

namespace ScoringEngine.Strategies;

public class RangeComparisonStrategy : IRuleStrategy
{
    public RuleResult Execute(ScoringRule rule, ExcelWorksheet worksheet)
    {
        var result = new RuleResult
        {
            RuleId = rule.RuleId,
            RuleType = nameof(RuleType.RangeComparison),
            Description = rule.Description,
            PointsMax = rule.Points
        };

        var rangeAddr  = rule.Parameters.GetValueOrDefault("RangeAddress", "A1:A10");
        var reference  = rule.Parameters.GetValueOrDefault("Reference", "").Trim();
        var tolStr     = rule.Parameters.GetValueOrDefault("Tolerance", "0");
        double.TryParse(tolStr, out double tolerance);

        try
        {
            var range = worksheet.Cells[rangeAddr];

            if (string.IsNullOrEmpty(reference))
            {
                // No reference sheet – just verify the range is non-empty
                bool hasData = range.Any(c => c.Value != null);
                result.Passed       = hasData;
                result.ActualValue  = hasData ? $"Vùng {rangeAddr} có dữ liệu" : $"Vùng {rangeAddr} trống";
                result.ExpectedValue = $"Vùng {rangeAddr} không trống";
                result.PointsEarned = result.Passed ? rule.Points : 0;
                result.Details      = result.Passed
                    ? $"Tìm thấy dữ liệu trong {rangeAddr}"
                    : $"Vùng {rangeAddr} không có dữ liệu";
                return result;
            }

            // Compare with a reference worksheet in the same workbook
            var refSheet = worksheet.Workbook.Worksheets[reference];
            if (refSheet == null)
            {
                result.Details = $"Không tìm thấy sheet tham chiếu '{reference}' trong workbook";
                return result;
            }

            int total = 0, matched = 0;
            foreach (var cell in range)
            {
                var refCell    = refSheet.Cells[cell.Address];
                var actualStr  = cell.Value?.ToString()   ?? "";
                var expectStr  = refCell.Value?.ToString() ?? "";

                bool cellMatch;
                if (double.TryParse(actualStr, out double aNum) &&
                    double.TryParse(expectStr,  out double eNum))
                    cellMatch = Math.Abs(aNum - eNum) <= tolerance;
                else
                    cellMatch = actualStr.Equals(expectStr, StringComparison.OrdinalIgnoreCase);

                if (cellMatch) matched++;
                total++;
            }

            result.Passed        = total > 0 && matched == total;
            result.PointsEarned  = result.Passed ? rule.Points : 0;
            result.ActualValue   = $"{matched}/{total} ô khớp";
            result.ExpectedValue = $"Khớp hoàn toàn với sheet '{reference}'";
            result.Details       = result.Passed
                ? $"Vùng {rangeAddr} khớp hoàn toàn với '{reference}'"
                : $"Vùng {rangeAddr}: {matched}/{total} ô đúng";
        }
        catch (Exception ex)
        {
            result.Details = $"Lỗi so sánh vùng: {ex.Message}";
        }

        return result;
    }
}
