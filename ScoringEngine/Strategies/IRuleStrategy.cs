using CommonLibrary.Models;
using OfficeOpenXml;
using ScoringEngine.Models;

namespace ScoringEngine.Strategies;

public interface IRuleStrategy
{
    RuleResult Execute(ScoringRule rule, ExcelWorksheet worksheet);
}
