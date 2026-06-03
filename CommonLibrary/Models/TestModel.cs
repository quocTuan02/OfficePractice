using System;
using System.Collections.Generic;

namespace CommonLibrary.Models
{
    public class TestModel
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public int TotalTimeMinutes { get; set; }
        public List<TaskModel> Tasks { get; set; }
        public DateTime CreatedDate { get; set; }
        public string Version { get; set; }

        public TestModel()
        {
            Id = Guid.NewGuid().ToString("N");
            Name = string.Empty;
            Description = string.Empty;
            TotalTimeMinutes = 60;
            Tasks = new List<TaskModel>();
            CreatedDate = DateTime.Now;
            Version = "1.0";
        }

        public int TotalPoints
        {
            get
            {
                int total = 0;
                foreach (var task in Tasks)
                    total += task.Points;
                return total;
            }
        }
    }
}
