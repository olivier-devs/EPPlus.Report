namespace EPPlus.Report.Model;

/// <summary>
///     Specifies options that control the behavior of group rendering.
/// </summary>
public class GroupOptions
{
    /// <summary>
    ///     Gets or sets a value indicating whether grouped rows should be collapsed.
    /// </summary>
    public bool Collapse { get; set; }

    /// <summary>
    ///     Gets or sets the merge mode for group labels.
    /// </summary>
    public MergeMode MergeLabels { get; set; } = MergeMode.None;

    /// <summary>
    ///     Gets or sets the column index where group labels should be placed.
    /// </summary>
    public int PlaceToColumn { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether the group should include a header.
    /// </summary>
    public bool WithHeader { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether subtotal rows should be disabled.
    /// </summary>
    public bool DisableSubtotals { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether the group outline should be disabled.
    /// </summary>
    public bool DisableOutline { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether page breaks should be inserted between groups.
    /// </summary>
    public bool PageBreaks { get; set; }

    /// <summary>
    ///     Gets or sets the label text for total rows.
    /// </summary>
    public string TotalLabel { get; set; } = "Total";

    /// <summary>
    ///     Gets or sets the label text for grand total rows.
    /// </summary>
    public string GrandLabel { get; set; } = "Grand";

    /// <summary>
    ///     Gets or sets a value indicating whether summary rows should appear above detail rows.
    /// </summary>
    public bool SummaryAbove { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether the grand total row should be disabled.
    /// </summary>
    public bool DisableGrandTotal { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether groups should be sorted in descending order.
    /// </summary>
    public bool Descending { get; set; }
}