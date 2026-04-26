using System.Collections.Generic;

namespace EPPlus.Report.Model;

/// <summary>
///     Represents a loop directive such as <c>&lt;&lt;foreach Items&gt;&gt;</c> and its child nodes.
/// </summary>
public class LoopNode : TemplateNode
{
    /// <summary>
    ///     Gets or sets the name of the collection to iterate over.
    /// </summary>
    public string CollectionName { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the child nodes inside the loop block.
    /// </summary>
    public List<TemplateNode> Children { get; set; } = [];

    /// <summary>
    ///     Gets or sets the row where the loop block ends.
    /// </summary>
    public int EndRow { get; set; }

    /// <summary>
    ///     Gets or sets the conditional formatting rules associated with this loop block.
    /// </summary>
    public List<ConditionalFormattingRule> ConditionalFormattingRules { get; set; } = [];
}