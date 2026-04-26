using System.Collections.Generic;

namespace EPPlus.Report.Model
{
    public class LoopNode : TemplateNode
    {
        public string CollectionName { get; set; } = string.Empty;
        public List<TemplateNode> Children { get; set; } = new List<TemplateNode>();
        public int EndRow { get; set; }
        public List<ConditionalFormattingRule> ConditionalFormattingRules { get; set; } = new List<ConditionalFormattingRule>();
    }
}
