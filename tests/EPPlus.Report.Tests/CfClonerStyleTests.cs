using System.Drawing;
using System.Linq;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class CfClonerStyleTests
    {
        private static void SetupLicense()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void Render_LoopWithCF_PreservesFillAndFontStyles()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Value}}";
            sheet.Cells["A3"].Value = "<</foreach>>";

            // Add CF with green fill and white bold font
            var cf = sheet.ConditionalFormatting.AddGreaterThan("A2");
            cf.Formula = "50";
            cf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cf.Style.Fill.BackgroundColor.SetColor(Color.Green);
            cf.Style.Font.Bold = true;
            cf.Style.Font.Color.SetColor(Color.White);
            cf.Priority = 5;
            cf.StopIfTrue = true;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = new[] { new { Value = 100 }, new { Value = 200 } } } }, sheet);

            // Assert: CF rule exists
            Assert.Equal(1, sheet.ConditionalFormatting.Count);
            var renderedRule = sheet.ConditionalFormatting.First();

            // Assert: fill style preserved
            Assert.Equal(OfficeOpenXml.Style.ExcelFillStyle.Solid, renderedRule.Style.Fill.PatternType);
            Assert.Equal(Color.Green.ToArgb(), renderedRule.Style.Fill.BackgroundColor.Color?.ToArgb());

            // Assert: font style preserved (Bold and Color are on ExcelDxfFontBase)
            Assert.True(renderedRule.Style.Font.Bold);
            Assert.Equal(Color.White.ToArgb(), renderedRule.Style.Font.Color.Color?.ToArgb());

            // Assert: priority and StopIfTrue preserved
            Assert.Equal(5, renderedRule.Priority);
            Assert.True(renderedRule.StopIfTrue);

            // Assert: formula preserved (via IExcelConditionalFormattingWithFormula)
            var formulaRule = Assert.IsAssignableFrom<IExcelConditionalFormattingWithFormula>(renderedRule);
            Assert.Equal("50", formulaRule.Formula);

            // Assert: covers all data rows
            Assert.True(renderedRule.Address.Start.Row <= 2);
            Assert.True(renderedRule.Address.End.Row >= 3);
        }
    }
}