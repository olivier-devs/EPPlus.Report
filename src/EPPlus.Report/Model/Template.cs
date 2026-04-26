using System.Collections.Generic;

namespace EPPlus.Report.Model
{
    public class Template
    {
        public List<TemplateNode> Nodes { get; set; } = new List<TemplateNode>();
    }
}
