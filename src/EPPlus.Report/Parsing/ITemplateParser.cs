using EPPlus.Report.Model;
using OfficeOpenXml;

namespace EPPlus.Report.Parsing
{
    public interface ITemplateParser
    {
        Template Parse(ExcelWorksheet worksheet, TemplateErrors errors);
    }
}
