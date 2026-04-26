namespace EPPlus.Report
{
    public class TemplateGenerateResult
    {
        public TemplateGenerateResult(TemplateErrors parsingErrors, TemplateErrors renderingErrors, TemplateErrors warnings = null)
        {
            ParsingErrors = parsingErrors;
            RenderingErrors = renderingErrors;
            Warnings = warnings ?? new TemplateErrors();
        }

        public TemplateGenerateResult(TemplateErrors errors) : this(errors, new TemplateErrors()) { }

        public bool HasErrors => ParsingErrors.Count > 0 || RenderingErrors.Count > 0;
        public bool HasWarnings => Warnings.Count > 0;
        public TemplateErrors ParsingErrors { get; }
        public TemplateErrors RenderingErrors { get; }
        public TemplateErrors Warnings { get; }
    }
}
