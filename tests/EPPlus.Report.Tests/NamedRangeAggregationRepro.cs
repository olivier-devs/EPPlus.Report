using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;
using System.IO;
using System;

namespace EPPlus.Report.Tests
{
    public class NamedRangeAggregationRepro
    {
        [Fact]
        public void Parse_NamedRange_With_Count_Aggregation_Inside()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{item.Name}}";
            sheet.Cells["A2"].Value = "<<count Items>>";
            sheet.Names.Add("Items", sheet.Cells["A1:A2"]);

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            var template = parser.Parse(sheet, errors);

            Assert.Single(template.Nodes);
            var nrLoop = Assert.IsType<NamedRangeLoopNode>(template.Nodes[0]);
            // The aggregation node should be preserved as a child
            Assert.Contains(nrLoop.Children, n => n is AggregationNode agg && agg.AggregationType == "count" && agg.PropertyName == "Items");
        }

        [Fact]
        public void Render_NamedRange_With_Count_Aggregation_Inside()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{item.Name}}";
            sheet.Cells["A2"].Value = "<<count Items>>";
            sheet.Names.Add("Items", sheet.Cells["A1:A2"]);

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
                Variables = new System.Collections.Generic.Dictionary<string, object> { { "Items", items } }
            };
            renderer.Render(template, context, sheet);
        }

        [Fact]
        public void Render_NamedRange_With_ServiceRow_Outside_Range_Should_Still_Work()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var tempFile = Path.GetTempFileName() + ".xlsx";
            try
            {
                using (var package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("Test");
                    sheet.Cells["A1"].Value = "Product";
                    sheet.Cells["B1"].Value = "Qty";
                    sheet.Cells["C1"].Value = "Price";
                    sheet.Cells["D1"].Value = "Category";
                    sheet.Cells["E1"].Value = "Total";
                    sheet.Cells["A2"].Value = "{{item.Product}}";
                    sheet.Cells["B2"].Value = "{{item.Qty}}";
                    sheet.Cells["C2"].Value = "{{item.Price}}";
                    sheet.Cells["D2"].Value = "{{item.Category}}";
                    sheet.Cells["E2"].Formula = "=B2*C2";
                    sheet.Cells["A3"].Value = "<<sum>>";
                    sheet.Cells["B3"].Value = "<<sum>>";
                    sheet.Cells["D3"].Value = "<<counta>>";
                    sheet.Cells["E3"].Value = "<<sum>>";
                    sheet.Names.Add("Products", sheet.Cells["A1:E2"]);
                    package.SaveAs(new FileInfo(tempFile));
                }

                var engine = new TemplateEngine(tempFile);
                engine.AddVariable("Products", new[]
                {
                    new { Product = "A", Qty = 2, Price = 10m, Category = "X" },
                    new { Product = "B", Qty = 3, Price = 20m, Category = "Y" }
                });
                var result = engine.Generate();

                var outputFile = Path.GetTempFileName() + ".xlsx";
                try
                {
                    engine.SaveAs(outputFile);
                    using var outPackage = new ExcelPackage(new FileInfo(outputFile));
                    var outSheet = outPackage.Workbook.Worksheets[0];

                    // Verify data rows
                    Assert.Equal("A", outSheet.Cells["A2"].Value);
                    Assert.Equal(2d, Convert.ToDouble(outSheet.Cells["B2"].Value));
                    Assert.Equal(10m, Convert.ToDecimal(outSheet.Cells["C2"].Value));
                    Assert.Equal("X", outSheet.Cells["D2"].Value);

                    Assert.Equal("B", outSheet.Cells["A3"].Value);
                    Assert.Equal(3d, Convert.ToDouble(outSheet.Cells["B3"].Value));
                    Assert.Equal(20m, Convert.ToDecimal(outSheet.Cells["C3"].Value));
                    Assert.Equal("Y", outSheet.Cells["D3"].Value);

                    // Verify service row formulas
                    Assert.Equal("SUBTOTAL(9,A2:A3)", outSheet.Cells["A4"].Formula);
                    Assert.Equal("SUBTOTAL(9,B2:B3)", outSheet.Cells["B4"].Formula);
                    Assert.Equal("SUBTOTAL(3,D2:D3)", outSheet.Cells["D4"].Formula);

                    outPackage.Workbook.Calculate();
                    Assert.Equal(5m, Convert.ToDecimal(outSheet.Cells["B4"].Value));
                    Assert.Equal(2, Convert.ToInt32(outSheet.Cells["D4"].Value));
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
    }
}
