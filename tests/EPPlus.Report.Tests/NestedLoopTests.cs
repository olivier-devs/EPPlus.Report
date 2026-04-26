using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class NestedLoopTests
    {
        [Fact]
        public void Parse_NestedLoop_CreatesNestedStructure()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Categories>>";
            sheet.Cells["A2"].Value = "{{CategoryName}}";
            sheet.Cells["A3"].Value = "<<foreach Products>>";
            sheet.Cells["A4"].Value = "{{ProductName}}";
            sheet.Cells["A5"].Value = "<</foreach>>";
            sheet.Cells["A6"].Value = "<</foreach>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var outerLoop = Assert.IsType<EPPlus.Report.Model.LoopNode>(template.Nodes[0]);
            Assert.Equal("Categories", outerLoop.CollectionName);
            var innerLoop = Assert.IsType<EPPlus.Report.Model.LoopNode>(outerLoop.Children[1]);
            Assert.Equal("Products", innerLoop.CollectionName);
        }

        [Fact]
        public void Render_NestedLoop_RendersCorrectly()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Categories>>";
            sheet.Cells["A2"].Value = "{{CategoryName}}";
            sheet.Cells["A3"].Value = "<<foreach Products>>";
            sheet.Cells["A4"].Value = "{{ProductName}}";
            sheet.Cells["A5"].Value = "<</foreach>>";
            sheet.Cells["A6"].Value = "<</foreach>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var categories = new[]
            {
                new 
                { 
                    CategoryName = "Electronics", 
                    Products = new[] { new { ProductName = "TV" }, new { ProductName = "Phone" } }
                },
                new 
                { 
                    CategoryName = "Books", 
                    Products = new[] { new { ProductName = "Novel" } }
                }
            };
            
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Categories = categories } }, sheet);
            
            Assert.Equal("Electronics", sheet.Cells["A2"].Value);
            Assert.Equal("TV", sheet.Cells["A4"].Value);
            Assert.Equal("Phone", sheet.Cells["A5"].Value);
            Assert.Equal("Books", sheet.Cells["A7"].Value);
            Assert.Equal("Novel", sheet.Cells["A9"].Value);
        }
    }
}
