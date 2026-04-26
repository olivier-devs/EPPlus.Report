using EPPlus.Report.Model;
using OfficeOpenXml;

namespace EPPlus.Report.Rendering
{
    public interface ITemplateRenderer
    {
        void Render(Template template, RenderContext context, ExcelWorksheet worksheet);
    }
}
