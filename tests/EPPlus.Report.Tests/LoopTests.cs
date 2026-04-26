using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class LoopTests
    {
        [Fact]
        public void Parse_SimpleLoop_CreatesLoopNodeWithChildren()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</foreach>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            Assert.Single(template.Nodes);
            var loopNode = Assert.IsType<EPPlus.Report.Model.LoopNode>(template.Nodes[0]);
            Assert.Equal("Items", loopNode.CollectionName);
            Assert.Single(loopNode.Children);
            var childExpr = Assert.IsType<EPPlus.Report.Model.ExpressionNode>(loopNode.Children[0]);
            Assert.Equal("Name", childExpr.ExpressionPath);
        }

        [Fact]
        public void Render_Loop_DuplicatesRows()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</foreach>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var items = new[]
            {
                new { Name = "Item1" },
                new { Name = "Item2" },
                new { Name = "Item3" }
            };
            
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);
            
            Assert.Equal("Item1", sheet.Cells["A2"].Value);
            Assert.Equal("Item2", sheet.Cells["A3"].Value);
            Assert.Equal("Item3", sheet.Cells["A4"].Value);
        }
    }
}
