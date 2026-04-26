namespace EPPlus.Report;

/// <summary>
///     Specifies options that control the template generation process.
/// </summary>
public class GenerateOptions
{
    /// <summary>
    ///     Gets or sets a value indicating whether formulas in the workbook should be evaluated after rendering.
    /// </summary>
    public bool EvaluateFormulas { get; set; } = false;
}