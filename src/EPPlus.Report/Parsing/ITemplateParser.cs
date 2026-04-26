using EPPlus.Report.Model;
using OfficeOpenXml;

namespace EPPlus.Report.Parsing;

/// <summary>
///     Defines methods for parsing an Excel worksheet into a template AST.
/// </summary>
public interface ITemplateParser
{
    /// <summary>
    ///     Parses the specified worksheet and builds a <see cref="Template" /> representing the template structure.
    /// </summary>
    /// <param name="worksheet">The Excel worksheet to parse.</param>
    /// <param name="errors">A collection to populate with parsing errors.</param>
    /// <returns>The parsed <see cref="Template" />.</returns>
    Template Parse(ExcelWorksheet worksheet, TemplateErrors errors);
}