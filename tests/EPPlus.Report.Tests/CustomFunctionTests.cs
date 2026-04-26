using System.IO;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class CustomFunctionTests
    {
        public class TestAddress
        {
            public string City { get; set; } = "Paris";
        }

        public class TestPerson
        {
            public string Name { get; set; } = "Alice";
            public TestAddress Address { get; set; } = new TestAddress();
        }

        private static void SetupLicense()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void Render_FunctionUpper_ReturnsUpperCase()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Upper(Name)}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var evaluator = new ExpressionEvaluator();
            var renderer = new TemplateRenderer(evaluator);
            renderer.Render(template, new RenderContext { Current = new TestPerson() }, sheet);

            Assert.Equal("ALICE", sheet.Cells["A1"].Value);
        }

        [Fact]
        public void Render_FunctionOnNestedProperty_ReturnsTransformed()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Upper(Address.City)}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var evaluator = new ExpressionEvaluator();
            var renderer = new TemplateRenderer(evaluator);
            renderer.Render(template, new RenderContext { Current = new TestPerson() }, sheet);

            Assert.Equal("PARIS", sheet.Cells["A1"].Value);
        }

        [Fact]
        public void Render_UnknownFunction_CollectsRenderingError()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Unknown(Name)}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var errors = new TemplateErrors();
            var evaluator = new ExpressionEvaluator();
            var renderer = new TemplateRenderer(evaluator, errors);
            renderer.Render(template, new RenderContext { Current = new TestPerson() }, sheet);

            Assert.Single(errors);
            Assert.Equal(ErrorType.Evaluation, errors[0].Type);
        }

        [Fact]
        public void Render_FunctionOnNull_ReturnsNull()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Upper(Name)}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var evaluator = new ExpressionEvaluator();
            var renderer = new TemplateRenderer(evaluator);
            renderer.Render(template, new RenderContext { Current = new TestPerson { Name = null } }, sheet);

            Assert.Null(sheet.Cells["A1"].Value);
        }

        [Fact]
        public void Parse_FunctionSyntax_SetsFunctionName()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Upper(Name)}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var exprNode = Assert.IsType<ExpressionNode>(template.Nodes[0]);
            Assert.Equal("Upper", exprNode.FunctionName);
            Assert.Equal("Name", exprNode.ExpressionPath);
        }

        [Fact]
        public void Render_BuiltInTrim_ReturnsTrimmed()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Trim(Name)}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var evaluator = new ExpressionEvaluator();
            var renderer = new TemplateRenderer(evaluator);
            renderer.Render(template, new RenderContext { Current = new TestPerson { Name = "  Alice  " } }, sheet);

            Assert.Equal("Alice", sheet.Cells["A1"].Value);
        }

        [Fact]
        public void Render_FunctionOverwrite_LastWins()
        {
            SetupLicense();
            using var inputStream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Test");
                sheet.Cells["A1"].Value = "{{Upper(Name)}}";
                package.SaveAs(inputStream);
            }
            inputStream.Position = 0;

            var engine = new TemplateEngine(inputStream);
            engine.AddVariable(new TestPerson { Name = "Alice" });
            engine.RegisterFunction("Upper", x => "FIRST");
            engine.RegisterFunction("Upper", x => "SECOND");
            var result = engine.Generate();

            Assert.False(result.HasErrors);

            using var outputStream = new MemoryStream();
            engine.SaveAs(outputStream);
            outputStream.Position = 0;

            using var resultPackage = new ExcelPackage(outputStream);
            var resultSheet = resultPackage.Workbook.Worksheets[0];
            Assert.Equal("SECOND", resultSheet.Cells["A1"].Value);
        }

        [Fact]
        public void Render_SimpleProperty_StillWorks()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Name}}";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var evaluator = new ExpressionEvaluator();
            var renderer = new TemplateRenderer(evaluator);
            renderer.Render(template, new RenderContext { Current = new TestPerson() }, sheet);

            Assert.Equal("Alice", sheet.Cells["A1"].Value);
        }
    }
}
