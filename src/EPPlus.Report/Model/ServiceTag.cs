namespace EPPlus.Report.Model;

/// <summary>
///     Represents a service tag found in a named range service row, such as <c>&lt;&lt;sum&gt;&gt;</c>,
///     <c>&lt;&lt;count&gt;&gt;</c>, <c>&lt;&lt;avg&gt;&gt;</c>, <c>&lt;&lt;max&gt;&gt;</c>, etc.
///     Service tags generate SUBTOTAL formulas for dynamic recalculation in Excel.
/// </summary>
public class ServiceTag
{
    /// <summary>
    ///     Gets or sets the name of the service tag (e.g., "sum", "count", "avg", "max", "min", etc.).
    ///     Supported tags: sum, count, counta, avg, max, min, product, stddev, stddevp, var, varp.
    /// </summary>
    public string TagName { get; set; } = string.Empty;

    /// <summary>
    ///     Gets or sets the row where the service tag is located.
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    ///     Gets or sets the column where the service tag is located.
    /// </summary>
    public int Column { get; set; }
}