using System.Drawing;
using System.Linq;
using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class ConditionalFormattingTests
    {
        private static void SetupLicense()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [Fact]
        public void Render_LoopWithCF_CopiesToInsertedRows()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Value}}";
            sheet.Cells["A3"].Value = "<</foreach>>";

            // Add CF to the template data row: red if > 100
            var cf = sheet.ConditionalFormatting.AddExpression("A2");
            cf.Formula = "A2>100";
            cf.Style.Fill.BackgroundColor.Color = Color.Red;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var items = new[]
            {
                new { Value = 50 },
                new { Value = 150 },
                new { Value = 200 }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            // After rendering: data rows are A2=50, A3=150, A4=200
            // The CF rule should exist and cover all data rows
            Assert.Equal(1, sheet.ConditionalFormatting.Count);
            var rule = sheet.ConditionalFormatting.First();
            Assert.True(rule.Address.Start.Row <= 2, "CF should start at or before row 2");
            Assert.True(rule.Address.End.Row >= 4, "CF should end at or after row 4");
        }

        [Fact]
        public void Render_IfFalse_RemovesCF()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<if Show>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</if>>";

            // Add CF to the data row inside the if block
            var cf = sheet.ConditionalFormatting.AddExpression("A2");
            cf.Formula = "A2>100";
            cf.Style.Fill.BackgroundColor.Color = Color.Red;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Show = false, Name = "Hidden" } }, sheet);

            // Block was deleted, no CF rules should remain
            Assert.Equal(0, sheet.ConditionalFormatting.Count);
        }

        [Fact]
        public void Render_GroupWithSubtotal_ExtendsCF()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<group Items by Category>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<<sum Price>>";
            sheet.Cells["A4"].Value = "<</group>>";

            // Add CF to the data row template
            var cf = sheet.ConditionalFormatting.AddExpression("A2");
            cf.Formula = "A2>100";
            cf.Style.Fill.BackgroundColor.Color = Color.Red;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var items = new[]
            {
                new { Category = "A", Name = "Item1", Price = 10m },
                new { Category = "A", Name = "Item2", Price = 20m },
                new { Category = "B", Name = "Item3", Price = 30m }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            // CF should exist and cover data rows plus subtotal rows
            Assert.Equal(1, sheet.ConditionalFormatting.Count);
            var rule = sheet.ConditionalFormatting.First();
            Assert.True(rule.Address.Start.Row <= 2, "CF should start at or before row 2");
            // Should extend at least to the last subtotal row (row 6 or beyond)
            Assert.True(rule.Address.End.Row >= 6, "CF should end at or after row 6 to cover subtotals");
        }

        [Fact]
        public void Render_CFWithFormula_AdjustsReferences()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Value}}";
            sheet.Cells["A3"].Value = "<</foreach>>";

            // Add CF with formula
            var cf = sheet.ConditionalFormatting.AddExpression("A2");
            cf.Formula = "A2>100";
            cf.Style.Fill.BackgroundColor.Color = Color.Red;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var items = new[]
            {
                new { Value = 50 },
                new { Value = 150 }
            };

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            // Verify the formula is preserved in the reconciled rule
            Assert.Equal(1, sheet.ConditionalFormatting.Count);
            var rule = sheet.ConditionalFormatting.First();
            var exprRule = Assert.IsAssignableFrom<OfficeOpenXml.ConditionalFormatting.Contracts.IExcelConditionalFormattingWithFormula>(rule);
            Assert.Equal("A2>100", exprRule.Formula?.ToString());
            Assert.True(rule.Address.Start.Row <= 2, "CF should start at or before row 2");
            Assert.True(rule.Address.End.Row >= 3, "CF should end at or after row 3");
        }

        [Fact]
        public void Parse_BlockWithCF_AssociatesRules()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</foreach>>";

            // Add CF to the data row inside the loop block
            var cf = sheet.ConditionalFormatting.AddExpression("A2");
            cf.Formula = "A2>100";
            cf.Style.Fill.BackgroundColor.Color = Color.Red;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var loopNode = Assert.IsType<LoopNode>(template.Nodes[0]);
            Assert.Single(loopNode.ConditionalFormattingRules);
            Assert.NotNull(loopNode.ConditionalFormattingRules[0].Address);
            Assert.Contains("A2", loopNode.ConditionalFormattingRules[0].Address);
        }
    }
}
