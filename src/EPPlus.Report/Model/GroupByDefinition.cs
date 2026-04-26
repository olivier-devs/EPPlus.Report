namespace EPPlus.Report.Model;

/// <summary>
///     Defines a grouping criterion for named range loops, including the property path and sort direction.
/// </summary>
public class GroupByDefinition
{
    /// <summary>
    ///     Gets or sets the property path used to extract group keys from items.
    /// </summary>
    public string PropertyPath { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the column index where the group key is located.
    /// </summary>
    public int Column { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether the group should be sorted in descending order.
    /// </summary>
    public bool Descending { get; set; }

    /// <summary>
    ///     Gets or sets the options controlling the rendering of this group.
    /// </summary>
    public GroupOptions Options { get; set; } = new();
}