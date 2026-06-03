namespace ScoringEngine.Models;

public class RuleResult
{
    public string RuleId { get; set; } = "";
    public string RuleType { get; set; } = "";
    public string Description { get; set; } = "";
    public bool Passed { get; set; }
    public int PointsEarned { get; set; }
    public int PointsMax { get; set; }
    public string ActualValue { get; set; } = "";
    public string ExpectedValue { get; set; } = "";
    public string Details { get; set; } = "";
    public string StatusIcon => Passed ? "✓" : "✗";
}
