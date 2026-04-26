namespace EPPlus.Report.Model
{
    public class ExpressionNode : TemplateNode
    {
        public string ExpressionPath { get; set; } = string.Empty;
        public string FunctionName { get; set; } = string.Empty;
    }
}
