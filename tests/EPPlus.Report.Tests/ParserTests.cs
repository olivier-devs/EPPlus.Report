using System;
using EPPlus.Report.Parsing;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class ParserTests
    {
        [Fact]
        public void Parse_SingleExpression_CreatesExpressionNode()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Name}}";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            Assert.Single(template.Nodes);
            var exprNode = Assert.IsType<EPPlus.Report.Model.ExpressionNode>(template.Nodes[0]);
            Assert.Equal("Name", exprNode.ExpressionPath);
            Assert.Equal(1, exprNode.Row);
            Assert.Equal(1, exprNode.Column);
        }

        [Fact]
        public void Parse_TextCell_CreatesTextNode()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "Hello World";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            Assert.Single(template.Nodes);
            Assert.IsType<EPPlus.Report.Model.TextNode>(template.Nodes[0]);
        }

        [Fact]
        public void Parse_MultipleCells_CreatesMultipleNodes()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{Name}}";
            sheet.Cells["B1"].Value = "{{Price}}";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            Assert.Equal(2, template.Nodes.Count);
        }

        [Fact]
        public void Parse_EmptyWorksheet_ReturnsEmptyTemplate()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            Assert.Empty(template.Nodes);
        }

        [Fact]
        public void Parse_ExpressionWithWhitespace_TrimmedPath()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "{{ Name }}";
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var exprNode = Assert.IsType<EPPlus.Report.Model.ExpressionNode>(template.Nodes[0]);
            Assert.Equal("Name", exprNode.ExpressionPath);
        }

        [Fact]
        public void Parse_NullWorksheet_ThrowsArgumentNullException()
        {
            var parser = new TemplateParser();
            Assert.Throws<ArgumentNullException>(() => parser.Parse(null!, new TemplateErrors()));
        }

        [Fact]
        public void Parse_UnclosedBlock_CreatesErrorWithFullLocationDetails()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";

            var parser = new TemplateParser();
            var errors = new TemplateErrors();
            parser.Parse(sheet, errors);

            Assert.Single(errors);
            var error = errors[0];
            Assert.Equal(ErrorType.Parsing, error.Type);
            Assert.Equal("Sheet1", error.WorksheetName);
            Assert.Equal(1, error.Row);
            Assert.Equal(1, error.Column);
            Assert.Equal("<<foreach Items>>", error.Expression);
            Assert.Equal("Sheet1!A1", error.Location);
            Assert.Equal("foreach block 'Items' is not closed", error.Message);
        }
    }
}
