namespace EPPlus.Report.Model
{
    public class GroupByDefinition
    {
        public string PropertyPath { get; set; } = string.Empty;
        public int Column { get; set; }
        public bool Descending { get; set; }
        public GroupOptions Options { get; set; } = new GroupOptions();
    }
}
