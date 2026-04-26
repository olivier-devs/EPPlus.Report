using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class IfTests
    {
        [Fact]
        public void Parse_SimpleIf_CreatesIfNodeWithChildren()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<if ShowName>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</if>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            Assert.Single(template.Nodes);
            var ifNode = Assert.IsType<EPPlus.Report.Model.IfNode>(template.Nodes[0]);
            Assert.Equal("ShowName", ifNode.ConditionExpression);
            Assert.Single(ifNode.Children);
        }

        [Fact]
        public void Render_IfTrue_ShowsContent()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<if Show>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</if>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Show = true, Name = "Visible" } }, sheet);
            
            Assert.Equal("Visible", sheet.Cells["A2"].Value);
        }

        [Fact]
        public void Render_IfFalse_HidesContent()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<if Show>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</if>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Show = false, Name = "Hidden" } }, sheet);
            
            Assert.Null(sheet.Cells["A2"].Value);
        }
    }
}
