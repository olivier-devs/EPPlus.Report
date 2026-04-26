namespace EPPlus.Report.Model
{
    public class AggregationNode : TemplateNode
    {
        public string AggregationType { get; set; } = string.Empty; // "sum", "count"
        public string PropertyName { get; set; } = string.Empty;
    }
}
