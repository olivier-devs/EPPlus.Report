using OfficeOpenXml.ConditionalFormatting;

namespace EPPlus.Report.Model
{
    public class ConditionalFormattingRule
    {
        public string Address { get; set; } = string.Empty;
        public string Formula { get; set; } = string.Empty;
        public string Formula2 { get; set; } = string.Empty;
        public eExcelConditionalFormattingRuleType Type { get; set; }
        public int Priority { get; set; }
        public bool StopIfTrue { get; set; }
    }
}
