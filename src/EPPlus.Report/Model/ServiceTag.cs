namespace EPPlus.Report.Model;

/// <summary>
///     Represents a service tag found in a named range service row, such as <c>&lt;&lt;sum&gt;&gt;</c> or
///     <c>&lt;&lt;count&gt;&gt;</c>.
/// </summary>
public class ServiceTag
{
    /// <summary>
    ///     Gets or sets the name of the service tag (e.g., "sum" or "count").
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