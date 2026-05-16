using System;
using System.Collections.Generic;
using System.IO;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class TemplateVisibleIntegrationTests : IDisposable
    {
        private readonly string _testOutputPath;

        public class SecureInvoiceModel
        {
            [TemplateVisible]
            public string InvoiceNumber { get; set; }

            [TemplateVisible]
            public decimal Amount { get; set; }

            public string InternalNotes { get; set; }
        }

        public TemplateVisibleIntegrationTests()
        {
            _testOutputPath = Path.Combine(Path.GetTempPath(), $"EPPlusTest_{Guid.NewGuid()}.xlsx");
        }

        public void Dispose()
        {
            if (File.Exists(_testOutputPath))
            {
                File.Delete(_testOutputPath);
            }
        }

        [Fact]
        public void TemplateEngine_WithTemplateVisible_MarkedPropertiesAreAccessible()
        {
            // Create a simple template
            using (var templateStream = new MemoryStream())
            {
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells[1, 1].Value = "{{InvoiceNumber}}";
                    worksheet.Cells[1, 2].Value = "{{Amount}}";
                    package.SaveAs(templateStream);
                }

                templateStream.Position = 0;

                using (var engine = new TemplateEngine(templateStream))
                {
                    var invoice = new SecureInvoiceModel
                    {
                        InvoiceNumber = "INV-001",
                        Amount = 100.50m,
                        InternalNotes = "Secret notes"
                    };

                    engine.AddVariable(invoice);
                    var result = engine.Generate();

                    Assert.False(result.HasErrors);
                    engine.SaveAs(_testOutputPath);
                }

                // Verify the output
                using (var resultPackage = new OfficeOpenXml.ExcelPackage(new FileInfo(_testOutputPath)))
                {
                    var worksheet = resultPackage.Workbook.Worksheets[0];
                    Assert.Equal("INV-001", worksheet.Cells[1, 1].Value);
                    Assert.Equal(100.50m, Convert.ToDecimal(worksheet.Cells[1, 2].Value), 2);
                }
            }
        }

        [Fact]
        public void TemplateEngine_WithTemplateVisible_NonMarkedPropertiesReportError()
        {
            // Create a template that tries to access a non-marked property
            using (var templateStream = new MemoryStream())
            {
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells[1, 1].Value = "{{InternalNotes}}";
                    package.SaveAs(templateStream);
                }

                templateStream.Position = 0;

                using (var engine = new TemplateEngine(templateStream))
                {
                    var invoice = new SecureInvoiceModel
                    {
                        InvoiceNumber = "INV-001",
                        InternalNotes = "Secret notes"
                    };

                    engine.AddVariable(invoice);

                    // Should report an error because InternalNotes is not in allowlist
                    var result = engine.Generate();
                    Assert.True(result.HasErrors);
                }
            }
        }

        [Fact]
        public void TemplateEngine_NoTemplateVisible_AllPropertiesAccessible()
        {
            // Without TemplateVisible attributes, all properties should be accessible (default behavior)
            using (var templateStream = new MemoryStream())
            {
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells[1, 1].Value = "{{InvoiceNumber}}";
                    worksheet.Cells[1, 2].Value = "{{InternalNotes}}";
                    package.SaveAs(templateStream);
                }

                templateStream.Position = 0;

                using (var engine = new TemplateEngine(templateStream))
                {
                    // Use anonymous type without TemplateVisible
                    var data = new { InvoiceNumber = "INV-003", InternalNotes = "Notes" };
                    engine.AddVariable(data);
                    var result = engine.Generate();

                    Assert.False(result.HasErrors);
                    engine.SaveAs(_testOutputPath);
                }

                using (var resultPackage = new OfficeOpenXml.ExcelPackage(new FileInfo(_testOutputPath)))
                {
                    var worksheet = resultPackage.Workbook.Worksheets[0];
                    Assert.Equal("INV-003", worksheet.Cells[1, 1].Value);
                    Assert.Equal("Notes", worksheet.Cells[1, 2].Value);
                }
            }
        }

        [Fact]
        public void TemplateEngine_AllowProperty_AddsToAllowlist()
        {
            // AllowProperty should add properties to the allowlist
            using (var templateStream = new MemoryStream())
            {
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells[1, 1].Value = "{{InvoiceNumber}}";
                    package.SaveAs(templateStream);
                }

                templateStream.Position = 0;

                using (var engine = new TemplateEngine(templateStream))
                {
                    var invoice = new SecureInvoiceModel
                    {
                        InvoiceNumber = "INV-004",
                        Amount = 200m,
                        InternalNotes = "Secret"
                    };

                    engine.AddVariable(invoice);
                    engine.AllowProperty("InvoiceNumber");
                    var result = engine.Generate();

                    Assert.False(result.HasErrors);
                    engine.SaveAs(_testOutputPath);
                }

                using (var resultPackage = new OfficeOpenXml.ExcelPackage(new FileInfo(_testOutputPath)))
                {
                    var worksheet = resultPackage.Workbook.Worksheets[0];
                    Assert.Equal("INV-004", worksheet.Cells[1, 1].Value);
                }
            }
        }

        [Fact]
        public void TemplateEngine_AllowProperty_MergesWithTemplateVisible()
        {
            // AllowProperty entries should merge with [TemplateVisible] entries, not overwrite
            using (var templateStream = new MemoryStream())
            {
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    worksheet.Cells[1, 1].Value = "{{InvoiceNumber}}";
                    worksheet.Cells[1, 2].Value = "{{Amount}}";
                    package.SaveAs(templateStream);
                }

                templateStream.Position = 0;

                using (var engine = new TemplateEngine(templateStream))
                {
                    var invoice = new SecureInvoiceModel
                    {
                        InvoiceNumber = "INV-005",
                        Amount = 300m,
                        InternalNotes = "Secret"
                    };

                    engine.AddVariable(invoice);
                    // AllowProperty adds "InvoiceNumber" — [TemplateVisible] also adds "InvoiceNumber" and "Amount"
                    // The merge should keep all three
                    engine.AllowProperty("InvoiceNumber");
                    var result = engine.Generate();

                    Assert.False(result.HasErrors);
                    engine.SaveAs(_testOutputPath);
                }

                using (var resultPackage = new OfficeOpenXml.ExcelPackage(new FileInfo(_testOutputPath)))
                {
                    var worksheet = resultPackage.Workbook.Worksheets[0];
                    Assert.Equal("INV-005", worksheet.Cells[1, 1].Value);
                    Assert.Equal(300m, Convert.ToDecimal(worksheet.Cells[1, 2].Value));
                }
            }
        }

        [Fact]
        public void TemplateEngine_AllowProperty_ThrowsOnEmptyPath()
        {
            using (var templateStream = new MemoryStream())
            {
                using (var package = new OfficeOpenXml.ExcelPackage())
                {
                    package.Workbook.Worksheets.Add("Sheet1");
                    package.SaveAs(templateStream);
                }

                templateStream.Position = 0;

                using (var engine = new TemplateEngine(templateStream))
                {
                    Assert.Throws<ArgumentException>(() => engine.AllowProperty(""));
                    Assert.Throws<ArgumentException>(() => engine.AllowProperty("  "));
                    Assert.Throws<ArgumentException>(() => engine.AllowProperty(null));
                }
            }
        }
    }
}