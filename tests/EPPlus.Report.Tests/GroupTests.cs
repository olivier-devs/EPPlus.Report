using System;
using System.Collections.Generic;
using System.Linq;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class GroupTests
    {
        private static void SetupLicense()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void Parse_GroupBlock_CreatesGroupNode()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Category>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</group>>";

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            Assert.Single(template.Nodes);
            var groupNode = Assert.IsType<GroupNode>(template.Nodes[0]);
            Assert.Equal("Items", groupNode.CollectionName);
            Assert.Single(groupNode.GroupByPaths);
            Assert.Equal("Category", groupNode.GroupByPaths[0]);
            Assert.Single(groupNode.Children);
        }

        [Fact]
        public void Parse_GroupBlockMultiLevel_ParsesMultiplePaths()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Country, City>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</group>>";

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var groupNode = Assert.IsType<GroupNode>(template.Nodes[0]);
            Assert.Equal(2, groupNode.GroupByPaths.Count);
            Assert.Equal("Country", groupNode.GroupByPaths[0]);
            Assert.Equal("City", groupNode.GroupByPaths[1]);
        }

        [Fact]
        public void Parse_GroupBlockWithDesc_ParsesDescending()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Category desc>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</group>>";

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var groupNode = Assert.IsType<GroupNode>(template.Nodes[0]);
            Assert.Single(groupNode.GroupByPaths);
            Assert.Equal("Category", groupNode.GroupByPaths[0]);
        }

        [Fact]
        public void Parse_GroupBlockWithSubtotal_DetectsSubtotalTemplate()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Category>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<<sum Price>>";
            sheet.Cells["A4"].Value = "<</group>>";

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var groupNode = Assert.IsType<GroupNode>(template.Nodes[0]);
            Assert.Single(groupNode.Children); // Only {{Name}}
            Assert.Single(groupNode.SubtotalTemplate); // <<sum Price>>
            var aggNode = Assert.IsType<AggregationNode>(groupNode.SubtotalTemplate[0]);
            Assert.Equal("sum", aggNode.AggregationType);
        }

        [Fact]
        public void Parse_NamedRangeWithGroupTag_ParsesGroupByDefinitions()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["B1"].Value = "Category";
            sheet.Cells["A2"].Value = "{{item.Name}}";
            sheet.Cells["B2"].Value = "{{item.Category}}";
            sheet.Cells["A3"].Value = "<<group Category>>";
            sheet.Cells["B3"].Value = "<<sum>>";
            sheet.Names.Add("Items", sheet.Cells["A1:B3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var nrLoop = Assert.IsType<NamedRangeLoopNode>(template.Nodes[0]);
            Assert.Single(nrLoop.GroupByDefinitions);
            Assert.Equal("Category", nrLoop.GroupByDefinitions[0].PropertyPath);
            Assert.Equal(1, nrLoop.GroupByDefinitions[0].Column); // Column A
        }

        [Fact]
        public void GroupRenderer_SortAndGroup_GroupsItemsByKey()
        {
            var items = new[]
            {
                new { Category = "A", Name = "Item1" },
                new { Category = "B", Name = "Item2" },
                new { Category = "A", Name = "Item3" }
            };

            var evaluator = new ExpressionEvaluator();
            var result = GroupRenderer.SortAndGroup(items.Cast<object>().ToList(), new List<string> { "Category" }, evaluator);

            Assert.Equal(2, result.Count);
            Assert.Equal("A", result[0].Key[0]);
            Assert.Equal(2, result[0].Items.Count);
            Assert.Equal("B", result[1].Key[0]);
            Assert.Single(result[1].Items);
        }

        [Fact]
        public void Render_GroupBlock_GroupsItems()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Category>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</group>>";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var items = new[]
            {
                new { Category = "B", Name = "Item2" },
                new { Category = "A", Name = "Item1" },
                new { Category = "A", Name = "Item3" }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            // After grouping by Category ascending: A items first, then B
            Assert.Equal("Item1", sheet.Cells["A2"].Value);
            Assert.Equal("Item3", sheet.Cells["A3"].Value);
            Assert.Equal("Item2", sheet.Cells["A4"].Value);
        }

        [Fact]
        public void Render_GroupBlockWithSubtotal_CalculatesGroupTotals()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Category>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<<sum Price>>";
            sheet.Cells["A4"].Value = "<</group>>";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var items = new[]
            {
                new { Category = "A", Name = "Item1", Price = 10m },
                new { Category = "A", Name = "Item2", Price = 20m },
                new { Category = "B", Name = "Item3", Price = 30m }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            Assert.Equal("Item1", sheet.Cells["A2"].Value);
            Assert.Equal("Item2", sheet.Cells["A3"].Value);
            Assert.Equal(30m, sheet.Cells["A4"].Value); // Group A subtotal
            Assert.Equal("Item3", sheet.Cells["A5"].Value);
            Assert.Equal(30m, sheet.Cells["A6"].Value); // Group B subtotal
        }

        [Fact]
        public void Render_GroupBlockWithGrandTotal_InsertsGrandTotal()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Category>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<<sum Price>>";
            sheet.Cells["A4"].Value = "<</group>>";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var items = new[]
            {
                new { Category = "A", Name = "Item1", Price = 10m },
                new { Category = "B", Name = "Item2", Price = 20m }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            Assert.Equal(10m, sheet.Cells["A3"].Value); // Group A subtotal
            Assert.Equal(20m, sheet.Cells["A5"].Value); // Group B subtotal
            Assert.Equal(30m, sheet.Cells["A6"].Value); // Grand total
        }

        [Fact]
        public void Render_NamedRangeWithGroup_ServiceRowGrouping()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Category}}";
            sheet.Cells["B2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<group Category>>";
            sheet.Cells["B3"].Value = "<<sum>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:B3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[]
            {
                new { Category = "B", Amount = 20m },
                new { Category = "A", Amount = 10m },
                new { Category = "A", Amount = 30m }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var ctx = new RenderContext
            {
                Current = null,
                Variables = new Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, ctx, sheet);

            package.Workbook.Calculate();

            // Header at A1, data starts at A2
            // Group A: 10, 30 → subtotal 40
            // Group B: 20 → subtotal 20
            Assert.Equal("A", sheet.Cells["A2"].Value);
            Assert.Equal(10m, sheet.Cells["B2"].Value);
            Assert.Equal("A", sheet.Cells["A3"].Value);
            Assert.Equal(30m, sheet.Cells["B3"].Value);
            Assert.Equal(40m, Convert.ToDecimal(sheet.Cells["B4"].Value)); // Group A subtotal
            Assert.Equal("B", sheet.Cells["A5"].Value);
            Assert.Equal(20m, Convert.ToDecimal(sheet.Cells["B5"].Value));
            Assert.Equal(20m, Convert.ToDecimal(sheet.Cells["B6"].Value)); // Group B subtotal
        }

        [Fact]
        public void Render_GroupBlockEmpty_DeletesBlock()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Category>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</group>>";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = new object[0] } }, sheet);

            Assert.Null(sheet.Cells["A1"].Value);
            Assert.Null(sheet.Cells["A2"].Value);
            Assert.Null(sheet.Cells["A3"].Value);
        }

        [Fact]
        public void Render_GroupBlockSingleGroup_RendersSubtotal()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Category>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<<sum Price>>";
            sheet.Cells["A4"].Value = "<</group>>";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var items = new[]
            {
                new { Category = "A", Name = "Item1", Price = 10m },
                new { Category = "A", Name = "Item2", Price = 20m }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            Assert.Equal("Item1", sheet.Cells["A2"].Value);
            Assert.Equal("Item2", sheet.Cells["A3"].Value);
            Assert.Equal(30m, sheet.Cells["A4"].Value); // Subtotal for single group
            Assert.Equal(30m, sheet.Cells["A5"].Value); // Grand total
        }
    }
}
