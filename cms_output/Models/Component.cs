using System;
using System.Collections.Generic;

namespace CMS.Models
{
    public enum ComponentStatus { Pending, Approved }
    public enum ComponentRating { STD, AECQ }
    public enum ComponentProcess { SMD, THT }

    public class Component
    {
        public int Id { get; set; }
        public string FekaNo { get; set; } = "";
        public string Category { get; set; } = "";
        public string Description { get; set; } = "";
        public string PartNumber { get; set; } = "";
        public string Manufacturer { get; set; } = "";
        public string Footprint { get; set; } = "";
        public ComponentRating Rating { get; set; } = ComponentRating.STD;
        public ComponentProcess Process { get; set; } = ComponentProcess.SMD;
        public string LibraryRef { get; set; } = "";
        public string FootprintRef { get; set; } = "";
        public ComponentStatus Status { get; set; } = ComponentStatus.Pending;
        public bool IsCustom { get; set; } = false;
        public Dictionary<string, string> ExtraFields { get; set; } = new Dictionary<string, string>();
    }
}
