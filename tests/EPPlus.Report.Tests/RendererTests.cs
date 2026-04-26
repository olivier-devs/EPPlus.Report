using System;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class RendererTests
    {
        public class TestProduct
        {
            public string Name { get; set; } = "Product";
            public decimal Price { get; set; } = 10.99m;
        }

        [Fact]
        public void Render_SingleExpression_ReplacesValue()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Name}}";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new TestProduct { Name = "Widget" } }, sheet);
            
            Assert.Equal("Widget", sheet.Cells["A1"].Value);
        }

        [Fact]
        public void Render_WithoutErrorCollector_ThrowsOnMissingProperty()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "{{Missing}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var renderer = new TemplateRenderer(new ExpressionEvaluator()); // no error collector
            var context = new RenderContext { Current = new { Name = "Test" } };

            Assert.Throws<PropertyNotFoundException>(() => renderer.Render(template, context, sheet));
        }

        [Fact]
        public void Render_PropertyNotFound_CollectsWarning()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "{{Missing}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var errors = new TemplateErrors();
            var warnings = new TemplateErrors();
            var renderer = new TemplateRenderer(new ExpressionEvaluator(), errors, null, warnings);
            var context = new RenderContext { Current = new { Name = "Test" } };
            renderer.Render(template, context, sheet);

            Assert.Empty(errors);
            Assert.Single(warnings);
            var warning = warnings[0];
            Assert.Equal(ErrorType.Warning, warning.Type);
            Assert.Equal("Sheet1", warning.WorksheetName);
            Assert.Equal(1, warning.Row);
            Assert.Equal(1, warning.Column);
            Assert.Equal("Missing", warning.Expression);
            Assert.Equal("Sheet1!A1", warning.Location);
            Assert.Equal("A1", warning.CellAddress);
            Assert.Contains("Missing", warning.Message);
            Assert.Null(sheet.Cells["A1"].Value);
        }

        [Fact]
        public void Render_NullReference_CollectsRenderingError()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "{{Child.Name}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var errors = new TemplateErrors();
            var renderer = new TemplateRenderer(new ExpressionEvaluator(), errors);
            var context = new RenderContext { Current = new { Child = (object)null } };
            renderer.Render(template, context, sheet);

            Assert.Single(errors);
            Assert.Equal(ErrorType.Evaluation, errors[0].Type);
        }

        [Fact]
        public void Render_MultipleMissingProperties_CollectsWarnings()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "{{Missing1}}";
            sheet.Cells["B1"].Value = "{{Missing2}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var errors = new TemplateErrors();
            var warnings = new TemplateErrors();
            var renderer = new TemplateRenderer(new ExpressionEvaluator(), errors, null, warnings);
            var context = new RenderContext { Current = new { Name = "Test" } };
            renderer.Render(template, context, sheet);

            Assert.Empty(errors);
            Assert.Equal(2, warnings.Count);
            Assert.Contains(warnings, e => e.Expression == "Missing1");
            Assert.Contains(warnings, e => e.Expression == "Missing2");
            Assert.Null(sheet.Cells["A1"].Value);
            Assert.Null(sheet.Cells["B1"].Value);
        }
    }
}
