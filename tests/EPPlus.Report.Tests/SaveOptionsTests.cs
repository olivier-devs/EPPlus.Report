using OfficeOpenXml;
using System;
using System.IO;
using Xunit;

namespace EPPlus.Report.Tests;

public class SaveOptionsTests
{
    private static string CreateSimpleTemplate()
    {
        var tempFile = Path.GetTempFileName() + ".xlsx";
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "{{Name}}";
            package.SaveAs(new FileInfo(tempFile));
        }

        return tempFile;
    }

    [Fact]
    public void SaveAs_WithPassword_SavesEncryptedFile()
    {
        // Arrange
        var templateFile = CreateSimpleTemplate();
        var outputFile = Path.GetTempFileName() + ".xlsx";
        try
        {
            var engine = new TemplateEngine(templateFile);
            engine.AddVariable(new { Name = "EncryptedTest" });
            engine.Generate();

            var saveOptions = new SaveOptions { Password = "p@ss" };

            // Act
            engine.SaveAs(outputFile, saveOptions);

            // Assert — opening without password should fail
            var ex = Assert.Throws<Exception>(() =>
            {
                using var pkg = new ExcelPackage(new FileInfo(outputFile));
                // Accessing workbook forces decryption attempt
                _ = pkg.Workbook;
            });
            Assert.Contains("password", ex.Message.ToLowerInvariant());

            // Assert — opening with correct password should succeed
            using (var pkg = new ExcelPackage(new FileInfo(outputFile), "p@ss"))
            {
                var sheet = pkg.Workbook.Worksheets[0];
                Assert.Equal("EncryptedTest", sheet.Cells["A1"].Value);
            }
        }
        finally
        {
            File.Delete(templateFile);
            File.Delete(outputFile);
        }
    }

    [Fact]
    public void SaveAs_FileInfo_WithPassword_SavesEncryptedFile()
    {
        // Arrange
        var templateFile = CreateSimpleTemplate();
        var outputFile = Path.GetTempFileName() + ".xlsx";
        try
        {
            var engine = new TemplateEngine(templateFile);
            engine.AddVariable(new { Name = "FileInfoEncrypted" });
            engine.Generate();

            var saveOptions = new SaveOptions { Password = "s3cret" };
            var fileInfo = new FileInfo(outputFile);

            // Act
            engine.SaveAs(fileInfo, saveOptions);

            // Assert — opening without password should fail
            var ex = Assert.Throws<Exception>(() =>
            {
                using var pkg = new ExcelPackage(fileInfo);
                _ = pkg.Workbook;
            });
            Assert.Contains("password", ex.Message.ToLowerInvariant());

            // Assert — opening with correct password should succeed
            using (var pkg = new ExcelPackage(fileInfo, "s3cret"))
            {
                var sheet = pkg.Workbook.Worksheets[0];
                Assert.Equal("FileInfoEncrypted", sheet.Cells["A1"].Value);
            }
        }
        finally
        {
            File.Delete(templateFile);
            File.Delete(outputFile);
        }
    }

    [Fact]
    public void SaveAs_Stream_WithPassword_ThrowsNotSupportedException()
    {
        // Arrange
        var templateFile = CreateSimpleTemplate();
        try
        {
            var engine = new TemplateEngine(templateFile);
            engine.AddVariable(new { Name = "StreamTest" });
            engine.Generate();

            var saveOptions = new SaveOptions { Password = "p@ss" };

            using var stream = new MemoryStream();

            // Act & Assert
            var ex = Assert.Throws<NotSupportedException>(() => engine.SaveAs(stream, saveOptions));
            Assert.Contains("Password-protected workbooks cannot be saved to a Stream", ex.Message);
        }
        finally
        {
            File.Delete(templateFile);
        }
    }

    [Fact]
    public void SaveAs_WithNullPassword_SavesWithoutEncryption()
    {
        // Arrange
        var templateFile = CreateSimpleTemplate();
        var outputFile = Path.GetTempFileName() + ".xlsx";
        try
        {
            var engine = new TemplateEngine(templateFile);
            engine.AddVariable(new { Name = "NoEncryption" });
            engine.Generate();

            var saveOptions = new SaveOptions { Password = null };

            // Act
            engine.SaveAs(outputFile, saveOptions);

            // Assert — file should be readable without a password
            using (var pkg = new ExcelPackage(new FileInfo(outputFile)))
            {
                var sheet = pkg.Workbook.Worksheets[0];
                Assert.Equal("NoEncryption", sheet.Cells["A1"].Value);
            }
        }
        finally
        {
            File.Delete(templateFile);
            File.Delete(outputFile);
        }
    }

    [Fact]
    public void SaveAs_WithEmptyPassword_SavesWithoutEncryption()
    {
        // Arrange
        var templateFile = CreateSimpleTemplate();
        var outputFile = Path.GetTempFileName() + ".xlsx";
        try
        {
            var engine = new TemplateEngine(templateFile);
            engine.AddVariable(new { Name = "EmptyPassword" });
            engine.Generate();

            var saveOptions = new SaveOptions { Password = "" };

            // Act
            engine.SaveAs(outputFile, saveOptions);

            // Assert — file should be readable without a password
            using (var pkg = new ExcelPackage(new FileInfo(outputFile)))
            {
                var sheet = pkg.Workbook.Worksheets[0];
                Assert.Equal("EmptyPassword", sheet.Cells["A1"].Value);
            }
        }
        finally
        {
            File.Delete(templateFile);
            File.Delete(outputFile);
        }
    }

    [Fact]
    public void SaveAs_WithWhitespacePassword_SavesWithoutEncryption()
    {
        // Arrange
        var templateFile = CreateSimpleTemplate();
        var outputFile = Path.GetTempFileName() + ".xlsx";
        try
        {
            var engine = new TemplateEngine(templateFile);
            engine.AddVariable(new { Name = "WhitespacePassword" });
            engine.Generate();

            var saveOptions = new SaveOptions { Password = "   " };

            // Act
            engine.SaveAs(outputFile, saveOptions);

            // Assert — file should be readable without a password
            using (var pkg = new ExcelPackage(new FileInfo(outputFile)))
            {
                var sheet = pkg.Workbook.Worksheets[0];
                Assert.Equal("WhitespacePassword", sheet.Cells["A1"].Value);
            }
        }
        finally
        {
            File.Delete(templateFile);
            File.Delete(outputFile);
        }
    }

    [Fact]
    public void SaveAs_WithPasswordAndFormulaEvaluation_SavesEncryptedFile()
    {
        // Arrange
        var templateFile = Path.GetTempFileName() + ".xlsx";
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "{{Value}}";
            sheet.Cells["B1"].Formula = "A1*2";
            package.SaveAs(new FileInfo(templateFile));
        }

        var outputFile = Path.GetTempFileName() + ".xlsx";
        try
        {
            var engine = new TemplateEngine(templateFile);
            engine.AddVariable(new { Value = 21 });
            engine.Generate();

            var saveOptions = new SaveOptions
            {
                Password = "f0rmula!",
                EvaluateFormulasBeforeSave = true
            };

            // Act
            engine.SaveAs(outputFile, saveOptions);

            // Assert — file should be encrypted and formulas evaluated
            using (var pkg = new ExcelPackage(new FileInfo(outputFile), "f0rmula!"))
            {
                var sheet = pkg.Workbook.Worksheets[0];
                Assert.Equal(21d, sheet.Cells["A1"].Value);
                // Formula should have been evaluated before save: 21*2 = 42
                Assert.Equal(42d, sheet.Cells["B1"].Value);
            }
        }
        finally
        {
            File.Delete(templateFile);
            File.Delete(outputFile);
        }
    }
}