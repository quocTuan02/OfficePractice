using System;

namespace CommonLibrary.Models
{
    public class TemplateInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public OfficeType OfficeType { get; set; }
        public string FilePath { get; set; }
        public DateTime CreatedDate { get; set; }
        public string Version { get; set; }

        public TemplateInfo()
        {
            Id = Guid.NewGuid().ToString("N");
            Name = string.Empty;
            Description = string.Empty;
            FilePath = string.Empty;
            Version = "1.0";
            CreatedDate = DateTime.Now;
        }
    }
}
