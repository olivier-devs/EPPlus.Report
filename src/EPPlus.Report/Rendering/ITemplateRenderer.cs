using EPPlus.Report.Model;
using OfficeOpenXml;

namespace EPPlus.Report.Rendering;

/// <summary>
///     Defines methods for rendering a parsed template into an Excel worksheet using a data context.
/// </summary>
public interface ITemplateRenderer
{
    /// <summary>
    ///     Renders the specified template into the worksheet using the provided context.
    /// </summary>
    /// <param name="template">The parsed template to render.</param>
    /// <param name="context">The data context for evaluating expressions.</param>
    /// <param name="worksheet">The target Excel worksheet.</param>
    void Render(Template template, RenderContext context, ExcelWorksheet worksheet);
}