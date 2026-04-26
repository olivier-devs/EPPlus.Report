using System.Collections.Generic;

namespace EPPlus.Report.Model;

/// <summary>
///     Represents a conditional directive such as <c>&lt;&lt;if Condition&gt;&gt;</c> and its child nodes.
/// </summary>
public class IfNode : TemplateNode
{
    /// <summary>
    ///     Gets or sets the condition expression to evaluate.
    /// </summary>
    public string ConditionExpression { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the child nodes inside the conditional block.
    /// </summary>
    public List<TemplateNode> Children { get; set; } = [];

    /// <summary>
    ///     Gets or sets the row where the conditional block ends.
    /// </summary>
    public int EndRow { get; set; }

    /// <summary>
    ///     Gets or sets the conditional formatting rules associated with this block.
    /// </summary>
    public List<ConditionalFormattingRule> ConditionalFormattingRules { get; set; } = [];
}