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
    public class CfClonerIconSetTests
    {
        private static void SetupLicense() { ExcelPackage.LicenseContext = LicenseContext.NonCommercial; }

        [Fact]
        public void Render_LoopWithThreeIconSet_PreservesSettings()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Value}}";
            sheet.Cells["A3"].Value = "<</foreach>>";

            var cf = sheet.ConditionalFormatting.AddThreeIconSet("A2", eExcelconditionalFormatting3IconsSetType.TrafficLights1);
            cf.ShowValue = false;
            cf.Reverse = true;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = new[] { new { Value = 10 }, new { Value = 50 }, new { Value = 90 } } } }, sheet);

            Assert.Equal(1, sheet.ConditionalFormatting.Count);
            var iconSet = sheet.ConditionalFormatting.First().As.ThreeIconSet;
            Assert.NotNull(iconSet);
            Assert.False(iconSet.ShowValue);
            Assert.True(iconSet.Reverse);
        }
    }
}