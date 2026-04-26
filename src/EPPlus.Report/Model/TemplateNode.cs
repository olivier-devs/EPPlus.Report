namespace EPPlus.Report.Model;

/// <summary>
///     Base class for all nodes in the template AST.
/// </summary>
public abstract class TemplateNode
{
    /// <summary>
    ///     Gets or sets the row number of this node in the worksheet.
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    ///     Gets or sets the column number of this node in the worksheet.
    /// </summary>
    public int Column { get; set; }

    /// <summary>
    ///     Gets or sets the raw text content of the cell from which this node was parsed.
    /// </summary>
    public string RawContent { get; set; } = string.Empty;
}