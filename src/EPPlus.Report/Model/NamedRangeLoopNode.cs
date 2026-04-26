using System.Collections.Generic;

namespace EPPlus.Report.Model
{
    public class NamedRangeLoopNode : LoopNode
    {
        public string RangeName { get; set; } = string.Empty;
        public bool IsHorizontal { get; set; }
        public int ServiceRowCount { get; set; }
        public List<ServiceTag> ServiceTags { get; set; } = new List<ServiceTag>();
        public int EndColumn { get; set; }
        public int HeaderRowCount { get; set; }
        public List<GroupByDefinition> GroupByDefinitions { get; set; } = new List<GroupByDefinition>();
        public GroupOptions RangeGroupOptions { get; set; } = new GroupOptions();
    }
}
