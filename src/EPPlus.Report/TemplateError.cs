namespace EPPlus.Report;

/// <summary>
///     Represents a single error or warning encountered during template processing.
/// </summary>
public class TemplateError
{
    /// <summary>
    ///     Gets or sets the error message.
    /// </summary>
    public string Message { get; set; }

    /// <summary>
    ///     Gets or sets the cell address where the error occurred (e.g., "A1").
    /// </summary>
    public string CellAddress { get; set; }

    /// <summary>
    ///     Gets or sets the name of the worksheet where the error occurred.
    /// </summary>
    public string WorksheetName { get; set; }

    /// <summary>
    ///     Gets or sets the row number where the error occurred.
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    ///     Gets or sets the column number where the error occurred.
    /// </summary>
    public int Column { get; set; }

    /// <summary>
    ///     Gets or sets the template expression that caused the error.
    /// </summary>
    public string Expression { get; set; }

    /// <summary>
    ///     Gets or sets the type of the error.
    /// </summary>
    public ErrorType Type { get; set; }

    /// <summary>
    ///     Gets the full location of the error in the format "Worksheet!CellAddress".
    /// </summary>
    public string Location => $"{WorksheetName}!{CellAddress}";
}