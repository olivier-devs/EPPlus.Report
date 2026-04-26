namespace EPPlus.Report;

/// <summary>
///     Specifies options that control how the rendered workbook is saved.
/// </summary>
public class SaveOptions
{
    /// <summary>
    ///     Gets or sets a value indicating whether formulas should be evaluated before saving the workbook.
    /// </summary>
    public bool EvaluateFormulasBeforeSave { get; set; } = false;
}