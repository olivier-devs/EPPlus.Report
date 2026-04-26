namespace EPPlus.Report;

/// <summary>
///     Represents the result of a template generation operation, including any errors or warnings.
/// </summary>
public class TemplateGenerateResult
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="TemplateGenerateResult" /> class.
    /// </summary>
    /// <param name="parsingErrors">Errors encountered during template parsing.</param>
    /// <param name="renderingErrors">Errors encountered during template rendering.</param>
    /// <param name="warnings">Optional warnings generated during generation.</param>
    public TemplateGenerateResult(TemplateErrors parsingErrors, TemplateErrors renderingErrors,
        TemplateErrors warnings = null)
    {
        ParsingErrors = parsingErrors;
        RenderingErrors = renderingErrors;
        Warnings = warnings ?? [];
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="TemplateGenerateResult" /> class with only parsing errors.
    /// </summary>
    /// <param name="errors">Errors encountered during template parsing.</param>
    public TemplateGenerateResult(TemplateErrors errors) : this(errors, [])
    {
    }

    /// <summary>
    ///     Gets a value indicating whether any parsing or rendering errors occurred.
    /// </summary>
    public bool HasErrors => ParsingErrors.Count > 0 || RenderingErrors.Count > 0;

    /// <summary>
    ///     Gets a value indicating whether any warnings were generated.
    /// </summary>
    public bool HasWarnings => Warnings.Count > 0;

    /// <summary>
    ///     Gets the errors encountered during template parsing.
    /// </summary>
    public TemplateErrors ParsingErrors { get; }

    /// <summary>
    ///     Gets the errors encountered during template rendering.
    /// </summary>
    public TemplateErrors RenderingErrors { get; }

    /// <summary>
    ///     Gets the warnings generated during template generation.
    /// </summary>
    public TemplateErrors Warnings { get; }
}