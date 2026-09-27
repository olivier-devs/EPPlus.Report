using System.Collections.Generic;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class RootReferenceAllowlistTests
    {
        private static void SetupLicense()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private class SensitiveInvoice
        {
            public string Number { get; set; } = "INV-001";

            public string InternalNotes { get; set; } = "secret-internal-notes";
        }

        // --- Unit: ExpressionEvaluator.IsRootReferenceAllowed ---

        [Fact]
        public void IsRootReferenceAllowed_NoAllowlist_ComplexObjectAllowed()
        {
            var evaluator = new ExpressionEvaluator();

            Assert.True(evaluator.IsRootReferenceAllowed(new SensitiveInvoice()));
        }

        [Fact]
        public void IsRootReferenceAllowed_AllowlistActive_TerminalValuesAllowed()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Number" }
            };

            Assert.True(evaluator.IsRootReferenceAllowed(null));
            Assert.True(evaluator.IsRootReferenceAllowed("plain string"));
            Assert.True(evaluator.IsRootReferenceAllowed(42));
            Assert.True(evaluator.IsRootReferenceAllowed(System.DateTime.Today));
        }

        [Fact]
        public void IsRootReferenceAllowed_AllowlistActive_ComplexObjectDenied()
        {
            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Number" }
            };

            Assert.False(evaluator.IsRootReferenceAllowed(new SensitiveInvoice()));
        }

        // --- Integration: named variable root reference (the renderer bypass) ---

        [Fact]
        public void Render_RootVariableReferenceWithAllowlist_DeniedAndNotDumped()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{invoice}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "invoice.Number" }
            };
            var renderingErrors = new TemplateErrors();
            var renderer = new TemplateRenderer(evaluator, renderingErrors);

            var invoice = new SensitiveInvoice();
            var context = new RenderContext
            {
                Current = null,
                Variables = new Dictionary<string, object> { ["invoice"] = invoice }
            };

            renderer.Render(template, context, sheet);

            Assert.NotEmpty(renderingErrors);
            Assert.All(renderingErrors, e => Assert.Equal(ErrorType.Evaluation, e.Type));
            Assert.Contains(renderingErrors, e => e.Expression == "invoice");
            Assert.Null(sheet.Cells["A1"].Value);
        }

        [Fact]
        public void Render_RootVariableReferenceWithoutAllowlist_ObjectDumpedUnchanged()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{invoice}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var evaluator = new ExpressionEvaluator();
            var renderer = new TemplateRenderer(evaluator);

            var invoice = new SensitiveInvoice();
            var context = new RenderContext
            {
                Current = null,
                Variables = new Dictionary<string, object> { ["invoice"] = invoice }
            };

            renderer.Render(template, context, sheet);

            Assert.Same(invoice, sheet.Cells["A1"].Value);
        }

        // --- Integration: NamedRange loop root references (item / index) ---

        [Fact]
        public void Render_NamedRangeItemRootReferenceWithAllowlist_Denied()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["B10"].Value = "Header";
            sheet.Cells["B11"].Value = "{{item}}";
            sheet.Names.Add("Items", sheet.Cells["B10:F12"]);

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Number" }
            };
            var renderingErrors = new TemplateErrors();
            var renderer = new TemplateRenderer(evaluator, renderingErrors);

            var items = new[] { new SensitiveInvoice(), new SensitiveInvoice() };
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            Assert.NotEmpty(renderingErrors);
            Assert.All(renderingErrors, e => Assert.Equal(ErrorType.Evaluation, e.Type));
            Assert.Contains(renderingErrors, e => e.Expression == "item");
        }

        [Fact]
        public void Render_NamedRangeItemRootReference_TerminalItemAllowed()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["B10"].Value = "Header";
            sheet.Cells["B11"].Value = "{{item}}";
            sheet.Names.Add("Items", sheet.Cells["B10:F12"]);

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Number" }
            };
            var renderingErrors = new TemplateErrors();
            var renderer = new TemplateRenderer(evaluator, renderingErrors);

            var items = new[] { "alpha", "beta" };
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            Assert.Empty(renderingErrors);
            Assert.Equal("alpha", sheet.Cells["B11"].Value);
        }

        [Fact]
        public void Render_NamedRangeIndexReferenceWithAllowlist_Allowed()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["B10"].Value = "Header";
            sheet.Cells["B11"].Value = "{{index}}";
            sheet.Names.Add("Items", sheet.Cells["B10:F12"]);

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var evaluator = new ExpressionEvaluator
            {
                AllowedProperties = new HashSet<string> { "Number" }
            };
            var renderingErrors = new TemplateErrors();
            var renderer = new TemplateRenderer(evaluator, renderingErrors);

            var items = new[] { new SensitiveInvoice(), new SensitiveInvoice() };
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            Assert.Empty(renderingErrors);
            Assert.NotNull(sheet.Cells["B11"].Value);
        }
    }
}
