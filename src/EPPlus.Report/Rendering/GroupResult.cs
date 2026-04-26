using System.Collections.Generic;

namespace EPPlus.Report.Rendering;

/// <summary>
///     Represents a single group resulting from a sort-and-group operation.
/// </summary>
public class GroupResult
{
    /// <summary>
    ///     Gets or sets the key values that define this group.
    /// </summary>
    public List<object> Key { get; set; }

    /// <summary>
    ///     Gets or sets the items belonging to this group.
    /// </summary>
    public List<object> Items { get; set; }
}