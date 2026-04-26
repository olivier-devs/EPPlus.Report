using System.Collections.Generic;

namespace EPPlus.Report.Model
{
    public class GroupNode : LoopNode
    {
        public List<string> GroupByPaths { get; set; } = new List<string>();
        public GroupOptions Options { get; set; } = new GroupOptions();
        public List<TemplateNode> SubtotalTemplate { get; set; } = new List<TemplateNode>();
    }
}
