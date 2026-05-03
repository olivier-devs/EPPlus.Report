using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using EPPlus.Report.Evaluation;
using OfficeOpenXml;
using System;
using System.IO;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class NamedRangeTests
    {
        private static void SetupLicense()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void Generate_NamedRangeVertical_RendersItems()
        {
            SetupLicense();
            var tempFile = Path.GetTempFileName() + ".xlsx";
            try
            {
                using (var package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("Test");
                    sheet.Cells["A1"].Value = "Order No";
                    sheet.Cells["B1"].Value = "Amount";
                    sheet.Cells["A2"].Value = "{{item.OrderNo}}";
                    sheet.Cells["B2"].Value = "{{item.Amount}}";
                    sheet.Cells["A3"].Value = "<<sum>>";
                    sheet.Names.Add("Orders", sheet.Cells["A1:B3"]);
                    package.SaveAs(new FileInfo(tempFile));
                }

                var engine = new TemplateEngine(tempFile);
                engine.AddVariable("Orders", new[]
                {
                    new { OrderNo = 100, Amount = 50m },
                    new { OrderNo = 101, Amount = 75m }
                });
                var result = engine.Generate();

                Assert.False(result.HasErrors);

                var outputFile = Path.GetTempFileName() + ".xlsx";
                try
                {
                    engine.SaveAs(outputFile);
                    using var outPackage = new ExcelPackage(new FileInfo(outputFile));
                    var outSheet = outPackage.Workbook.Worksheets[0];
                    outPackage.Workbook.Calculate();
                    Assert.Equal(100d, Convert.ToDouble(outSheet.Cells["A2"].Value));
                    Assert.Equal(50m, Convert.ToDecimal(outSheet.Cells["B2"].Value));
                    Assert.Equal(101d, Convert.ToDouble(outSheet.Cells["A3"].Value));
                    Assert.Equal(75m, Convert.ToDecimal(outSheet.Cells["B3"].Value));
                }
                finally
                {
                    File.Delete(outputFile);
                }
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void Parse_NamedRange_CreatesNamedRangeLoopNode()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{item.Name}}";
            sheet.Cells["A2"].Value = "{{item.Value}}";
            sheet.Names.Add("Items", sheet.Cells["A1:A2"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            Assert.Single(template.Nodes);
            var nrLoop = Assert.IsType<NamedRangeLoopNode>(template.Nodes[0]);
            Assert.Equal("Items", nrLoop.CollectionName);
            Assert.Equal("Items", nrLoop.RangeName);
            Assert.False(nrLoop.IsHorizontal);
        }

        [Fact]
        public void Parse_NamedRangeWithIndex_ParsesCorrectly()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{index}}";
            sheet.Cells["A2"].Value = "{{item.Name}}";
            sheet.Names.Add("Items", sheet.Cells["A1:A2"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var nrLoop = Assert.IsType<NamedRangeLoopNode>(template.Nodes[0]);
            Assert.Contains(nrLoop.Children, n => n is ExpressionNode expr && expr.ExpressionPath == "index");
            Assert.Contains(nrLoop.Children, n => n is ExpressionNode expr && expr.ExpressionPath == "item.Name");
        }

        [Fact]
        public void Parse_NamedRangeOverlapsForeach_ForeachWins()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</foreach>>";
            sheet.Cells["B1"].Value = "{{item.Value}}";
            sheet.Names.Add("Items", sheet.Cells["B1:B2"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            // Should have the explicit LoopNode, not the NamedRangeLoopNode
            var loopNodes = template.Nodes.FindAll(n => n is LoopNode);
            Assert.Single(loopNodes);
            Assert.IsType<LoopNode>(loopNodes[0]);
            Assert.DoesNotContain(template.Nodes, n => n is NamedRangeLoopNode);
        }

        [Fact]
        public void Parse_NamedRangeAndClassicForeach_BothWork()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            
            // Classic foreach at A1:A3
            sheet.Cells["A1"].Value = "<<foreach ClassicItems>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</foreach>>";
            
            // Named range at C1:D2 (no overlap)
            sheet.Cells["C1"].Value = "{{item.Value}}";
            sheet.Names.Add("NamedItems", sheet.Cells["C1:D2"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            Assert.Contains(template.Nodes, n => n is LoopNode && !(n is NamedRangeLoopNode));
            Assert.Contains(template.Nodes, n => n is NamedRangeLoopNode);
        }

        [Fact]
        public void Render_NamedRangeWithServiceTag_GeneratesSubtotal()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<sum>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[]
            {
                new { Amount = 10m },
                new { Amount = 20m },
                new { Amount = 30m }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            // Verify data rows
            Assert.Equal(10m, sheet.Cells["A2"].Value);
            Assert.Equal(20m, sheet.Cells["A3"].Value);
            Assert.Equal(30m, sheet.Cells["A4"].Value);

            // Verify service row has SUBTOTAL formula
            Assert.Equal("SUBTOTAL(9,A2:A4)", sheet.Cells["A5"].Formula);

            package.Workbook.Calculate();
            Assert.Equal(60m, Convert.ToDecimal(sheet.Cells["A5"].Value));
        }

        [Fact]
        public void Render_NamedRangeWithSumTag_InsertsSubtotal9Formula()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<sum>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[]
            {
                new { Amount = 10m },
                new { Amount = 20m }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(9,A2:A3)", sheet.Cells["A4"].Formula);
        }

        [Fact]
        public void Render_NamedRangeWithCountTag_InsertsSubtotal3Formula()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["A2"].Value = "{{item.Name}}";
            sheet.Cells["A3"].Value = "<<count>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[]
            {
                new { Name = "A" },
                new { Name = "B" },
                new { Name = "C" }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(3,A2:A4)", sheet.Cells["A5"].Formula);
        }

        [Fact]
        public void Render_NamedRangeWithSumTag_EvaluatesCorrectlyAfterCalculate()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<sum>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[]
            {
                new { Amount = 15m },
                new { Amount = 25m },
                new { Amount = 35m }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            package.Workbook.Calculate();
            Assert.Equal(75m, Convert.ToDecimal(sheet.Cells["A5"].Value));
        }

        [Fact]
        public void Render_NamedRangeWithAvgTag_GeneratesSubtotal1()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<avg>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[] { new { Amount = 10m }, new { Amount = 20m }, new { Amount = 30m } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(1,A2:A4)", sheet.Cells["A5"].Formula);

            package.Workbook.Calculate();
            Assert.Equal(20m, Convert.ToDecimal(sheet.Cells["A5"].Value));
        }

        [Fact]
        public void Render_NamedRangeWithMaxTag_GeneratesSubtotal4()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<max>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[] { new { Amount = 10m }, new { Amount = 30m }, new { Amount = 20m } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(4,A2:A4)", sheet.Cells["A5"].Formula);

            package.Workbook.Calculate();
            Assert.Equal(30m, Convert.ToDecimal(sheet.Cells["A5"].Value));
        }

        [Fact]
        public void Render_NamedRangeWithCountaTag_GeneratesSubtotal3()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Name";
            sheet.Cells["A2"].Value = "{{item.Name}}";
            sheet.Cells["A3"].Value = "<<counta>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[] { new { Name = "Alice" }, new { Name = "Bob" }, new { Name = "Charlie" } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(3,A2:A4)", sheet.Cells["A5"].Formula);

            package.Workbook.Calculate();
            Assert.Equal(3, Convert.ToInt32(sheet.Cells["A5"].Value));
        }

        [Fact]
        public void Render_NamedRangeWithMinTag_GeneratesSubtotal5()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<min>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[] { new { Amount = 10m }, new { Amount = 30m }, new { Amount = 20m } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(5,A2:A4)", sheet.Cells["A5"].Formula);

            package.Workbook.Calculate();
            Assert.Equal(10m, Convert.ToDecimal(sheet.Cells["A5"].Value));
        }

        [Fact]
        public void Render_NamedRangeWithProductTag_GeneratesSubtotal6()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<product>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[] { new { Amount = 2m }, new { Amount = 3m }, new { Amount = 4m } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(6,A2:A4)", sheet.Cells["A5"].Formula);

            package.Workbook.Calculate();
            Assert.Equal(24m, Convert.ToDecimal(sheet.Cells["A5"].Value)); // 2*3*4=24
        }

        [Fact]
        public void Render_NamedRangeWithStddevTag_GeneratesSubtotal7()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<stddev>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[] { new { Amount = 10m }, new { Amount = 20m }, new { Amount = 30m } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext { Current = null, Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } } };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(7,A2:A4)", sheet.Cells["A5"].Formula);
            package.Workbook.Calculate();
            var result = Convert.ToDouble(sheet.Cells["A5"].Value);
            Assert.InRange(result, 9.9, 10.1);
        }

        [Fact]
        public void Render_NamedRangeWithVarTag_GeneratesSubtotal10()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<var>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[] { new { Amount = 10m }, new { Amount = 20m }, new { Amount = 30m } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext
            {
                Current = null,
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } }
            };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(10,A2:A4)", sheet.Cells["A5"].Formula);
            package.Workbook.Calculate();
            var result = Convert.ToDouble(sheet.Cells["A5"].Value);
            Assert.InRange(result, 99.9, 100.1);
        }

        [Fact]
        public void Render_NamedRangeWithStddevpTag_GeneratesSubtotal8()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<stddevp>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[] { new { Amount = 10m }, new { Amount = 20m }, new { Amount = 30m } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext { Current = null, Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } } };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(8,A2:A4)", sheet.Cells["A5"].Formula);
            package.Workbook.Calculate();
            var result = Convert.ToDouble(sheet.Cells["A5"].Value);
            Assert.InRange(result, 8.1, 8.2);
        }

        [Fact]
        public void Render_NamedRangeWithVarpTag_GeneratesSubtotal11()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Amount";
            sheet.Cells["A2"].Value = "{{item.Amount}}";
            sheet.Cells["A3"].Value = "<<varp>>";
            sheet.Names.Add("Sales", sheet.Cells["A1:A3"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            var items = new[] { new { Amount = 10m }, new { Amount = 20m }, new { Amount = 30m } };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var context = new RenderContext { Current = null, Variables = new System.Collections.Generic.Dictionary<string, object> { { "Sales", items } } };
            renderer.Render(template, context, sheet);

            Assert.Equal("SUBTOTAL(11,A2:A4)", sheet.Cells["A5"].Formula);
            package.Workbook.Calculate();
            var result = Convert.ToDouble(sheet.Cells["A5"].Value);
            Assert.InRange(result, 66.6, 66.7);
        }
    }
}
