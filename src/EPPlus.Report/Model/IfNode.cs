using System.Collections.Generic;

namespace EPPlus.Report.Model
{
    public class IfNode : TemplateNode
    {
        public string ConditionExpression { get; set; } = string.Empty;
        public List<TemplateNode> Children { get; set; } = new List<TemplateNode>();
        public int EndRow { get; set; }
        public List<ConditionalFormattingRule> ConditionalFormattingRules { get; set; } = new List<ConditionalFormattingRule>();
    }
}
