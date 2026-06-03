namespace ScoringEngine.Models;

public class TaskScoreResult
{
    public string TaskId { get; set; } = "";
    public int TaskNumber { get; set; }
    public string TaskObjective { get; set; } = "";
    public List<RuleResult> RuleResults { get; set; } = new();
    public int TotalPoints => RuleResults.Sum(r => r.PointsEarned);
    public int MaxPoints => RuleResults.Sum(r => r.PointsMax);
    public bool AllPassed => RuleResults.Count > 0 && RuleResults.All(r => r.Passed);
    public string Summary => $"{TotalPoints}/{MaxPoints} điểm ({RuleResults.Count(r => r.Passed)}/{RuleResults.Count} rules đúng)";
}
