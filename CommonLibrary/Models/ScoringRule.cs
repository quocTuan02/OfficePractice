using System;
using System.Collections.Generic;

namespace CommonLibrary.Models
{
    public class ScoringRule
    {
        public string RuleId { get; set; }
        public string TaskId { get; set; }
        public RuleType RuleType { get; set; }
        public string Description { get; set; }
        public int Points { get; set; }
        public Dictionary<string, string> Parameters { get; set; }

        public ScoringRule()
        {
            RuleId = Guid.NewGuid().ToString("N");
            TaskId = string.Empty;
            Description = string.Empty;
            Parameters = new Dictionary<string, string>();
        }
    }
}
