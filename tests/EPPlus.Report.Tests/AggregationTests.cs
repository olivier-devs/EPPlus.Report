using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class AggregationTests
    {
        [Fact]
        public void Parse_Sum_CreatesAggregationNode()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<sum Price>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            Assert.Single(template.Nodes);
            var aggNode = Assert.IsType<EPPlus.Report.Model.AggregationNode>(template.Nodes[0]);
            Assert.Equal("sum", aggNode.AggregationType);
            Assert.Equal("Price", aggNode.PropertyName);
        }

        [Fact]
        public void Parse_Count_CreatesAggregationNode()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<count Items>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            Assert.Single(template.Nodes);
            var aggNode = Assert.IsType<EPPlus.Report.Model.AggregationNode>(template.Nodes[0]);
            Assert.Equal("count", aggNode.AggregationType);
            Assert.Equal("Items", aggNode.PropertyName);
        }

        [Fact]
        public void Parse_SumInsideLoop_ParsesAsChildOfLoop()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<<sum Price>>";
            sheet.Cells["A4"].Value = "<</foreach>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var loopNode = Assert.IsType<EPPlus.Report.Model.LoopNode>(template.Nodes[0]);
            Assert.Equal(2, loopNode.Children.Count);
            var aggNode = Assert.IsType<EPPlus.Report.Model.AggregationNode>(loopNode.Children[1]);
            Assert.Equal("sum", aggNode.AggregationType);
        }

        [Fact]
        public void Render_SumInsideLoop_CalculatesTotal()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<<sum Price>>";
            sheet.Cells["A4"].Value = "<</foreach>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var items = new[]
            {
                new { Name = "A", Price = 10m },
                new { Name = "B", Price = 20m },
                new { Name = "C", Price = 30m }
            };
            
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);
            
            // Debug: check other cells
            Assert.Null(sheet.Cells["A1"].Value); // start tag removed
            Assert.Equal("A", sheet.Cells["A2"].Value); // first item name
            
            // Sum should appear in each row (aggregating the entire collection)
            Assert.Equal(60m, sheet.Cells["A3"].Value);
        }

        [Fact]
        public void Render_CountInsideLoop_CalculatesCount()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<<count Items>>";
            sheet.Cells["A4"].Value = "<</foreach>>";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var items = new[]
            {
                new { Name = "A" },
                new { Name = "B" },
                new { Name = "C" }
            };
            
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);
            
            Assert.Equal(3, sheet.Cells["A3"].Value);
        }

        [Fact]
        public void Render_CountAtRootLevel_WithNamedVariable_CalculatesCount()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<count Entities>>";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var entities = new[] { new { Name = "A" }, new { Name = "B" } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Entities", entities } }
            };
            renderer.Render(template, context, sheet);

            Assert.Equal(2, sheet.Cells["A1"].Value);
        }

        [Fact]
        public void Render_CountAtRootLevel_WithCurrentObject_CalculatesCount()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<count Items>>";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var items = new[] { new { Name = "A" }, new { Name = "B" }, new { Name = "C" } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            Assert.Equal(3, sheet.Cells["A1"].Value);
        }

        [Fact]
        public void Render_Sum_WithNonNumericValues_IgnoresNonNumericValues()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<<sum Price>>";
            sheet.Cells["A4"].Value = "<</foreach>>";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            // Mix of numeric and non-numeric values
            var items = new[]
            {
                new { Name = "A", Price = (object)10m },
                new { Name = "B", Price = (object)"not a number" },
                new { Name = "C", Price = (object)30m }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            // Sum should only include numeric values (10 + 30 = 40)
            Assert.Equal(40m, sheet.Cells["A3"].Value);
        }
    }
}
