using System.Drawing;
using System.Linq;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class CfClonerDataBarTests
    {
        private static void SetupLicense() { ExcelPackage.LicenseContext = LicenseContext.NonCommercial; }

        [Fact]
        public void Render_LoopWithDataBar_PreservesColorAndBounds()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Value}}";
            sheet.Cells["A3"].Value = "<</foreach>>";

            var cf = sheet.ConditionalFormatting.AddDatabar("A2", Color.SteelBlue);
            cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Min;
            cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Max;
            cf.ShowValue = false;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = new[] { new { Value = 10 }, new { Value = 80 } } } }, sheet);

            Assert.Equal(1, sheet.ConditionalFormatting.Count);
            var dataBar = sheet.ConditionalFormatting.First().As.DataBar;
            Assert.NotNull(dataBar);
            Assert.Equal(Color.SteelBlue.ToArgb(), dataBar.Color.ToArgb());
            Assert.False(dataBar.ShowValue);
            Assert.Equal(eExcelConditionalFormattingValueObjectType.Min, dataBar.LowValue.Type);
            Assert.Equal(eExcelConditionalFormattingValueObjectType.Max, dataBar.HighValue.Type);
        }
    }
}