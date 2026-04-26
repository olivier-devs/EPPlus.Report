using System.Collections.Generic;

namespace EPPlus.Report.Model;

/// <summary>
///     Represents the parsed structure of an Excel template, containing a list of template nodes.
/// </summary>
public class Template
{
    /// <summary>
    ///     Gets or sets the top-level nodes that make up the template.
    /// </summary>
    public List<TemplateNode> Nodes { get; set; } = [];
}