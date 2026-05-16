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
    public class CfClonerColorScaleTests
    {
        private static void SetupLicense() { ExcelPackage.LicenseContext = LicenseContext.NonCommercial; }

        [Fact]
        public void Render_LoopWithThreeColorScale_PreservesColors()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Value}}";
            sheet.Cells["A3"].Value = "<</foreach>>";

            var cf = sheet.ConditionalFormatting.AddThreeColorScale("A2");
            cf.LowValue.Color = Color.FromArgb(255, 99, 190, 123);
            cf.LowValue.Type = eExcelConditionalFormattingValueObjectType.Min;
            cf.MiddleValue.Color = Color.FromArgb(255, 255, 235, 132);
            cf.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
            cf.MiddleValue.Value = 50;
            cf.HighValue.Color = Color.FromArgb(255, 248, 105, 107);
            cf.HighValue.Type = eExcelConditionalFormattingValueObjectType.Max;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = new[] { new { Value = 10 }, new { Value = 50 }, new { Value = 90 } } } }, sheet);

            Assert.Equal(1, sheet.ConditionalFormatting.Count);
            var scale = sheet.ConditionalFormatting.First().As.ThreeColorScale;
            Assert.NotNull(scale);
            Assert.Equal(Color.FromArgb(255, 99, 190, 123).ToArgb(), scale.LowValue.Color.ToArgb());
            Assert.Equal(eExcelConditionalFormattingValueObjectType.Min, scale.LowValue.Type);
            Assert.Equal(Color.FromArgb(255, 255, 235, 132).ToArgb(), scale.MiddleValue.Color.ToArgb());
            Assert.Equal(eExcelConditionalFormattingValueObjectType.Percent, scale.MiddleValue.Type);
            Assert.Equal(50, scale.MiddleValue.Value);
            Assert.Equal(Color.FromArgb(255, 248, 105, 107).ToArgb(), scale.HighValue.Color.ToArgb());
        }
    }
}