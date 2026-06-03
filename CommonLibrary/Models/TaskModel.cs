using System;
using System.Collections.Generic;

namespace CommonLibrary.Models
{
    public class TaskModel
    {
        public string TaskId { get; set; }
        public int Number { get; set; }
        public string Instruction { get; set; }
        public int Points { get; set; }
        public string TemplateFile { get; set; }
        public string Objective { get; set; }
        public List<ScoringRule> ScoringRules { get; set; }

        public TaskModel()
        {
            TaskId = Guid.NewGuid().ToString("N");
            Instruction = string.Empty;
            TemplateFile = string.Empty;
            Objective = string.Empty;
            ScoringRules = new List<ScoringRule>();
        }
    }
}
