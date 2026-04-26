namespace EPPlus.Report.Model;

/// <summary>
///     Defines how group labels should be merged during group rendering.
/// </summary>
public enum MergeMode
{
    /// <summary>
    ///     Group labels are not merged.
    /// </summary>
    None,

    /// <summary>
    ///     Merges group labels and clears duplicate cells.
    /// </summary>
    Merge1,

    /// <summary>
    ///     Merges group labels and clears duplicate cells (variant 2).
    /// </summary>
    Merge2,

    /// <summary>
    ///     Merges group labels without clearing duplicate cells.
    /// </summary>
    Merge3
}