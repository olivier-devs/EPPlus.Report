using System.Collections.Generic;

namespace EPPlus.Report.Model;

/// <summary>
///     Represents a loop node derived from an Excel named range, supporting headers, service rows, and grouping.
/// </summary>
public class NamedRangeLoopNode : LoopNode
{
    /// <summary>
    ///     Gets or sets the name of the Excel named range.
    /// </summary>
    public string RangeName { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets a value indicating whether the loop iterates horizontally.
    /// </summary>
    public bool IsHorizontal { get; set; }

    /// <summary>
    ///     Gets or sets the number of service rows at the end of the named range.
    /// </summary>
    public int ServiceRowCount { get; set; }

    /// <summary>
    ///     Gets or sets the service tags (such as sum or count) found in the service row.
    /// </summary>
    public List<ServiceTag> ServiceTags { get; set; } = [];

    /// <summary>
    ///     Gets or sets the column where the named range ends.
    /// </summary>
    public int EndColumn { get; set; }

    /// <summary>
    ///     Gets or sets the number of header rows at the start of the named range.
    /// </summary>
    public int HeaderRowCount { get; set; }

    /// <summary>
    ///     Gets or sets the group-by definitions for named range grouping.
    /// </summary>
    public List<GroupByDefinition> GroupByDefinitions { get; set; } = [];

    /// <summary>
    ///     Gets or sets the group options specific to the named range.
    /// </summary>
    public GroupOptions RangeGroupOptions { get; set; } = new();
}