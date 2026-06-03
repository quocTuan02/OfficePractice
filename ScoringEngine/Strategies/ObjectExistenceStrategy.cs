using CommonLibrary.Models;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using ScoringEngine.Models;

namespace ScoringEngine.Strategies;

public class ObjectExistenceStrategy : IRuleStrategy
{
    public RuleResult Execute(ScoringRule rule, ExcelWorksheet worksheet)
    {
        var result = new RuleResult
        {
            RuleId = rule.RuleId,
            RuleType = nameof(RuleType.ObjectExistence),
            Description = rule.Description,
            PointsMax = rule.Points
        };

        var objectType = rule.Parameters.GetValueOrDefault("ObjectType", "Chart");
        var objectName = rule.Parameters.GetValueOrDefault("ObjectName", "");
        result.ExpectedValue = $"{objectType}{(string.IsNullOrEmpty(objectName) ? " (bất kỳ)" : $" '{objectName}'")} tồn tại";

        try
        {
            bool exists = objectType switch
            {
                "Chart" => CheckDrawing<ExcelChart>(worksheet, objectName),
                "PivotTable" => CheckPivot(worksheet, objectName),
                "Table" => CheckTable(worksheet, objectName),
                "Image" => CheckImage(worksheet, objectName),
                "Shape" => CheckAnyDrawing(worksheet, objectName),
                _ => false
            };

            result.Passed = exists;
            result.PointsEarned = exists ? rule.Points : 0;
            result.ActualValue = exists ? $"{objectType} tồn tại" : $"{objectType} không tồn tại";
            result.Details = result.Passed
                ? $"Tìm thấy {objectType}{(string.IsNullOrEmpty(objectName) ? "" : $" '{objectName}'")}"
                : $"Không tìm thấy {objectType}{(string.IsNullOrEmpty(objectName) ? "" : $" '{objectName}'")}";
        }
        catch (Exception ex)
        {
            result.Details = $"Lỗi kiểm tra: {ex.Message}";
        }

        return result;
    }

    private static bool CheckDrawing<T>(ExcelWorksheet ws, string name) where T : ExcelDrawing
    {
        var items = ws.Drawings.OfType<T>();
        return string.IsNullOrEmpty(name)
            ? items.Any()
            : items.Any(d => d.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    private static bool CheckPivot(ExcelWorksheet ws, string name)
        => string.IsNullOrEmpty(name)
            ? ws.PivotTables.Count > 0
            : ws.PivotTables.Any(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    private static bool CheckTable(ExcelWorksheet ws, string name)
        => string.IsNullOrEmpty(name)
            ? ws.Tables.Count > 0
            : ws.Tables.Any(t => t.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

    private static bool CheckImage(ExcelWorksheet ws, string name)
    {
        var pics = ws.Drawings.Where(d => d.DrawingType == OfficeOpenXml.Drawing.eDrawingType.Picture);
        return string.IsNullOrEmpty(name)
            ? pics.Any()
            : pics.Any(d => d.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    private static bool CheckAnyDrawing(ExcelWorksheet ws, string name)
        => string.IsNullOrEmpty(name)
            ? ws.Drawings.Count > 0
            : ws.Drawings.Any(d => d.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
}
