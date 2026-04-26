namespace EPPlus.Report.Model
{
    public abstract class TemplateNode
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string RawContent { get; set; } = string.Empty;
    }
}
