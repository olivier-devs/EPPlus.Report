using BenchmarkDotNet.Attributes;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace EPPlus.Report.Benchmarks
{
    [MemoryDiagnoser]
    public class TemplateBenchmarks
    {
        public class Product
        {
            public string Name { get; set; }
            public decimal Price { get; set; }
            public int Quantity { get; set; }
            public string Category { get; set; }
        }

        public class Company
        {
            public string Name { get; set; }
            public Product[] Products { get; set; }
        }

        [GlobalSetup]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Benchmark]
        [Arguments(1000)]
        [Arguments(5000)]
        [Arguments(10000)]
        [Arguments(20000)]
        public void LoopRender(int itemCount)
        {
            var products = Enumerable.Range(0, itemCount)
                .Select(i => new Product
                {
                    Name = $"Product {i}",
                    Price = i * 1.5m,
                    Quantity = i,
                    Category = $"Category {i % 10}"
                })
                .ToArray();

            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Products>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["B2"].Value = "{{Price}}";
            sheet.Cells["C2"].Value = "{{Quantity}}";
            sheet.Cells["D2"].Value = "{{Category}}";
            sheet.Cells["A3"].Value = "<</foreach>>";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new Company { Name = "TestCorp", Products = products } }, sheet);
        }

        [Benchmark]
        public void SimpleExpressions()
        {
            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "{{Name}}";
                sheet.Cells["B1"].Value = "{{Price}}";
                sheet.Cells["C1"].Value = "{{Quantity}}";
                sheet.Cells["D1"].Value = "{{Category}}";
                package.SaveAs(stream);
            }
            stream.Position = 0;

            var engine = new TemplateEngine(stream);
            engine.AddVariable(new Product
            {
                Name = "Widget",
                Price = 99.99m,
                Quantity = 100,
                Category = "Electronics"
            });
            engine.Generate();
        }

        [Benchmark]
        [Arguments(1000)]
        [Arguments(5000)]
        [Arguments(10000)]
        public void GroupedLoopRender(int itemCount)
        {
            var products = Enumerable.Range(0, itemCount)
                .Select(i => new Product
                {
                    Name = $"Product {i}",
                    Price = i * 1.5m,
                    Quantity = i,
                    Category = $"Category {i % 10}"
                })
                .ToArray();

            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "<<group Products by Category>>";
                sheet.Cells["A2"].Value = "{{Category}}";
                sheet.Cells["A3"].Value = "{{Name}}";
                sheet.Cells["B3"].Value = "{{Price}}";
                sheet.Cells["A4"].Value = "<</group>>";
                package.SaveAs(stream);
            }
            stream.Position = 0;

            var engine = new TemplateEngine(stream);
            engine.AddVariable(new Company { Name = "TestCorp", Products = products });
            engine.Generate();
        }

        [Benchmark]
        [Arguments(1000)]
        [Arguments(5000)]
        [Arguments(10000)]
        public void NamedRangeLoopRender(int itemCount)
        {
            var products = Enumerable.Range(0, itemCount)
                .Select(i => new Product
                {
                    Name = $"Product {i}",
                    Price = i * 1.5m,
                    Quantity = i,
                    Category = $"Category {i % 10}"
                })
                .ToArray();

            using var stream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "Name";
                sheet.Cells["B1"].Value = "Price";
                sheet.Cells["A2"].Value = "{{item.Name}}";
                sheet.Cells["B2"].Value = "{{item.Price}}";
                sheet.Cells["A3"].Value = "<<sum Price>>";
                package.SaveAs(stream);
            }
            stream.Position = 0;

            var engine = new TemplateEngine(stream);
            engine.AddVariable(new Company { Name = "TestCorp", Products = products });
            engine.Generate();
        }
    }
}
