using EPPlus.Report.Model;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class TemplateEngineTests
    {
        [Fact]
        public void Render_SimpleTemplate_GeneratesOutput()
        {
            // Create temp template file
            var tempFile = Path.GetTempFileName() + ".xlsx";
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("Sheet1");
                    sheet.Cells["A1"].Value = "{{Name}}";
                    package.SaveAs(new FileInfo(tempFile));
                }

                var engine = new TemplateEngine(tempFile);
                engine.AddVariable(new { Name = "Test" });
                var result = engine.Generate();
                Assert.False(result.HasErrors);

                var outputFile = Path.GetTempFileName() + ".xlsx";
                try
                {
                    engine.SaveAs(outputFile);
                    using var resultPackage = new ExcelPackage(new FileInfo(outputFile));
                    var resultSheet = resultPackage.Workbook.Worksheets[0];
                    Assert.Equal("Test", resultSheet.Cells["A1"].Value);
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
        public void Generate_NamedVariable_ResolvesInTemplate()
        {
            var tempFile = Path.GetTempFileName() + ".xlsx";
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("Sheet1");
                    sheet.Cells["A1"].Value = "{{CompanyName}}";
                    package.SaveAs(new FileInfo(tempFile));
                }

                var engine = new TemplateEngine(tempFile);
                engine.AddVariable("CompanyName", "Acme Corp");
                var result = engine.Generate();

                Assert.False(result.HasErrors);

                var outputFile = Path.GetTempFileName() + ".xlsx";
                try
                {
                    engine.SaveAs(outputFile);
                    using var resultPackage = new ExcelPackage(new FileInfo(outputFile));
                    var resultSheet = resultPackage.Workbook.Worksheets[0];
                    Assert.Equal("Acme Corp", resultSheet.Cells["A1"].Value);
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
        public void Generate_WithStream_SavesToOutputStream()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var inputStream = new MemoryStream();
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = "{{Value}}";
                package.SaveAs(inputStream);
            }
            inputStream.Position = 0;

            var engine = new TemplateEngine(inputStream);
            engine.AddVariable(new { Value = 42 });
            var result = engine.Generate();

            Assert.False(result.HasErrors);

            using var outputStream = new MemoryStream();
            engine.SaveAs(outputStream);
            outputStream.Position = 0;

            using var resultPackage = new ExcelPackage(outputStream);
            var resultSheet = resultPackage.Workbook.Worksheets[0];
            Assert.Equal(42d, resultSheet.Cells["A1"].Value);
        }

        [Fact]
        public void Generate_UnclosedLoop_ReturnsError()
        {
            var tempFile = Path.GetTempFileName() + ".xlsx";
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("Sheet1");
                    sheet.Cells["A1"].Value = "<<foreach Items>>";
                    sheet.Cells["A2"].Value = "{{Name}}";
                    // missing <</foreach>>
                    package.SaveAs(new FileInfo(tempFile));
                }

                var engine = new TemplateEngine(tempFile);
                engine.AddVariable(new { Items = new[] { new { Name = "A" } } });
                var result = engine.Generate();

                Assert.True(result.HasErrors);
                Assert.Contains(result.ParsingErrors, e => e.Type == ErrorType.Parsing);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void Generate_PropertyNotFound_ReturnsWarning()
        {
            var tempFile = Path.GetTempFileName() + ".xlsx";
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("Sheet1");
                    sheet.Cells["A1"].Value = "{{MissingProperty}}";
                    package.SaveAs(new FileInfo(tempFile));
                }

                var engine = new TemplateEngine(tempFile);
                engine.AddVariable(new { Name = "Test" });
                var result = engine.Generate();

                Assert.False(result.HasErrors);
                Assert.Empty(result.RenderingErrors);
                Assert.True(result.HasWarnings);
                Assert.Single(result.Warnings);
                var warning = result.Warnings[0];
                Assert.Equal(ErrorType.Warning, warning.Type);
                Assert.Equal("Sheet1", warning.WorksheetName);
                Assert.Equal(1, warning.Row);
                Assert.Equal(1, warning.Column);
                Assert.Equal("MissingProperty", warning.Expression);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }

        [Fact]
        public void Dispose_DoesNotThrow()
        {
            var tempFile = Path.GetTempFileName() + ".xlsx";
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage())
                {
                    var sheet = package.Workbook.Worksheets.Add("Sheet1");
                    sheet.Cells["A1"].Value = "{{Name}}";
                    package.SaveAs(new FileInfo(tempFile));
                }

                var engine = new TemplateEngine(tempFile);
                var exception = Record.Exception(() => engine.Dispose());
                Assert.Null(exception);
            }
            finally
            {
                File.Delete(tempFile);
            }
        }
    }
}
