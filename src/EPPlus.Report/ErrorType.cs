namespace EPPlus.Report;

/// <summary>
///     Defines the types of errors that can occur during template processing.
/// </summary>
public enum ErrorType
{
    /// <summary>
    ///     An error occurred while parsing the template structure.
    /// </summary>
    Parsing,

    /// <summary>
    ///     An error occurred while evaluating a template expression.
    /// </summary>
    Evaluation,

    /// <summary>
    ///     An error occurred while rendering the template to the output.
    /// </summary>
    Rendering,

    /// <summary>
    ///     A non-fatal warning was generated during processing.
    /// </summary>
    Warning
}