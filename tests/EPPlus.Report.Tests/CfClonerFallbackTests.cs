using System.Linq;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class CfClonerFallbackTests
    {
        private static void SetupLicense() { ExcelPackage.LicenseContext = LicenseContext.NonCommercial; }

        [Fact]
        public void Render_LoopWithAboveAverage_DoesNotCrash_FallsBackToV1()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Value}}";
            sheet.Cells["A3"].Value = "<</foreach>>";

            // AboveAverage is a cell-style CF type - should be cloned properly
            var cf = sheet.ConditionalFormatting.AddAboveAverage("A2");

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = new[] { new { Value = 10 }, new { Value = 80 } } } }, sheet);

            // CF should exist and cover all data rows
            Assert.Equal(1, sheet.ConditionalFormatting.Count);
            var renderedRule = sheet.ConditionalFormatting.First();
            Assert.True(renderedRule.Address.Start.Row <= 2);
            Assert.True(renderedRule.Address.End.Row >= 3);
        }
    }
}