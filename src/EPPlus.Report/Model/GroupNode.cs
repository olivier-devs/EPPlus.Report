using System.Collections.Generic;

namespace EPPlus.Report.Model;

/// <summary>
///     Represents a group directive such as <c>&lt;&lt;group Items by Category&gt;&gt;</c> with support for subtotals and
///     grand totals.
/// </summary>
public class GroupNode : LoopNode
{
    /// <summary>
    ///     Gets or sets the property paths used to group items.
    /// </summary>
    public List<string> GroupByPaths { get; set; } = [];

    /// <summary>
    ///     Gets or sets the options controlling group rendering behavior.
    /// </summary>
    public GroupOptions Options { get; set; } = new();

    /// <summary>
    ///     Gets or sets the template nodes used to render subtotal and grand total rows.
    /// </summary>
    public List<TemplateNode> SubtotalTemplate { get; set; } = [];
}