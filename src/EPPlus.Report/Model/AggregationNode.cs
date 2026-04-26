namespace EPPlus.Report.Model;

/// <summary>
///     Represents an aggregation directive such as <c>&lt;&lt;sum Property&gt;&gt;</c> or
///     <c>&lt;&lt;count Items&gt;&gt;</c>.
/// </summary>
public class AggregationNode : TemplateNode
{
    /// <summary>
    ///     Gets or sets the type of aggregation, such as "sum" or "count".
    /// </summary>
    public string AggregationType { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the property name or collection name to aggregate.
    /// </summary>
    public string PropertyName { get; set; } = string.Empty;
}