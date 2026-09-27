using System;
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
    public class ConditionalFormattingReconciliationTests
    {
        private static void SetupLicense()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private static void BuildNamedRangeTemplate(ExcelWorksheet sheet)
        {
            sheet.Cells["B10"].Value = "Header";
            sheet.Cells["B11"].Value = "{{item.Name}}";
            sheet.Cells["D11"].Value = "{{item.Amount}}";
            sheet.Names.Add("Items", sheet.Cells["B10:F12"]);
        }

        private static object[] Items =>
            new object[]
            {
                new { Name = "A", Amount = 50m },
                new { Name = "B", Amount = 150m },
                new { Name = "C", Amount = 200m }
            };

        // Test 8 : association NamedRange + rectangle du bloc
        [Fact]
        public void Parse_NamedRangeWithCF_AssociatesRuleWithBlockRectangle()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            BuildNamedRangeTemplate(sheet);
            var cf = sheet.ConditionalFormatting.AddExpression("D10:D12");
            cf.Formula = "D11>100";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());

            var node = Assert.IsType<NamedRangeLoopNode>(template.Nodes[0]);
            Assert.Equal(2, node.Column);
            Assert.Equal(6, node.EndColumn);
            var rule = Assert.Single(node.ConditionalFormattingRules);
            Assert.Equal(10, rule.TemplateStartRow);
            Assert.Equal(12, rule.TemplateEndRow);
            Assert.Equal(2, rule.TemplateStartColumn);
            Assert.Equal(6, rule.TemplateEndColumn);
        }

        // Test 1 : CF colonne unique — colonnes préservées, lignes étendues
        [Fact]
        public void Render_NamedRangeLoopWithColumnCF_PreservesColumnAndExtendsRows()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            BuildNamedRangeTemplate(sheet);
            var cf = sheet.ConditionalFormatting.AddExpression("D10:D12");
            cf.Formula = "D11>100";
            cf.Style.Fill.BackgroundColor.Color = Color.Red;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items } }, sheet);

            Assert.Equal(1, sheet.ConditionalFormatting.Count);
            var rule = sheet.ConditionalFormatting.First();
            Assert.Equal(4, rule.Address.Start.Column);
            Assert.Equal(4, rule.Address.End.Column);
            Assert.True(rule.Address.Start.Row <= 10, "CF should start at or before the block start row");
            Assert.True(rule.Address.End.Row >= 16, "CF should cover the final block rows");
            Assert.Equal(Color.Red.ToArgb(), rule.Style.Fill.BackgroundColor.Color?.ToArgb());
        }

        // Test 2 : CF multi-colonnes — lignes étendues, colonnes conservées
        [Fact]
        public void Render_NamedRangeLoopWithMultiColumnCF_ExtendsRowsKeepsColumns()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            BuildNamedRangeTemplate(sheet);
            var cf = sheet.ConditionalFormatting.AddExpression("D10:F12");
            cf.Formula = "D11>100";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items } }, sheet);

            var rule = Assert.Single(sheet.ConditionalFormatting);
            Assert.Equal(4, rule.Address.Start.Column);
            Assert.Equal(6, rule.Address.End.Column);
            Assert.True(rule.Address.End.Row >= 16, "CF should cover the final block rows");
        }

        // Test 3 : CF pleine colonne — reste D:D
        [Fact]
        public void Render_NamedRangeLoopWithFullColumnCF_KeepsFullColumnAddress()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            BuildNamedRangeTemplate(sheet);
            var cf = sheet.ConditionalFormatting.AddExpression("D:D");
            cf.Formula = "D11>100";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items } }, sheet);

            var rule = Assert.Single(sheet.ConditionalFormatting);
            Assert.Equal("D:D", rule.Address.Address);
        }

        // Test 4 : CF extérieure — non associée, non recréée
        [Fact]
        public void Render_NamedRangeLoopWithOutsideCF_DoesNotAssociateOrRecreate()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            BuildNamedRangeTemplate(sheet);
            var cf = sheet.ConditionalFormatting.AddExpression("H10:H12");
            cf.Formula = "H11>100";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var node = Assert.IsType<NamedRangeLoopNode>(template.Nodes[0]);
            Assert.Empty(node.ConditionalFormattingRules);

            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items } }, sheet);

            var rule = Assert.Single(sheet.ConditionalFormatting);
            Assert.Equal(8, rule.Address.Start.Column);
            Assert.Equal(8, rule.Address.End.Column);
        }

        // Test 5 : collection vide — CF du bloc supprimée
        [Fact]
        public void Render_NamedRangeLoopEmptyCollection_RemovesBlockCF()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            BuildNamedRangeTemplate(sheet);
            var cf = sheet.ConditionalFormatting.AddExpression("D10:D12");
            cf.Formula = "D11>100";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = Array.Empty<object>() } }, sheet);

            Assert.Equal(0, sheet.ConditionalFormatting.Count);
        }

        // Test 6 : collection vide — CF interne supprimée, CF externe conservée
        [Fact]
        public void Render_NamedRangeLoopEmptyCollection_KeepsOutsideCF()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            BuildNamedRangeTemplate(sheet);
            var cfIn = sheet.ConditionalFormatting.AddExpression("D10:D12");
            cfIn.Formula = "D11>100";
            var cfOut = sheet.ConditionalFormatting.AddExpression("H10:H20");
            cfOut.Formula = "H11>100";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items = Array.Empty<object>() } }, sheet);

            var rule = Assert.Single(sheet.ConditionalFormatting);
            Assert.Equal(8, rule.Address.Start.Column);
        }

        // Test 7 : règle déjà ajustée — pas de duplication (compromis boucle classique)
        [Fact]
        public void Render_ClassicLoopWithColumnCF_AlreadyAdjustedRuleNotDuplicated()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            sheet.Cells["B10"].Value = "<<foreach Items>>";
            sheet.Cells["B11"].Value = "{{Value}}";
            sheet.Cells["B12"].Value = "<</foreach>>";
            var cf = sheet.ConditionalFormatting.AddExpression("D10:D12");
            cf.Formula = "D11>100";
            cf.Style.Fill.BackgroundColor.Color = Color.Red;

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            var items = new[] { new { Value = 50 }, new { Value = 150 }, new { Value = 200 } };
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);

            var rule = Assert.Single(sheet.ConditionalFormatting);
            Assert.Equal(4, rule.Address.Start.Column);
            Assert.Equal(4, rule.Address.End.Column);
            Assert.True(rule.Address.End.Row >= 14, "CF should cover the final block rows");
            Assert.Equal(Color.Red.ToArgb(), rule.Style.Fill.BackgroundColor.Color?.ToArgb());
        }

        // Cas 5 : CF partiellement extérieure — coordonnées réelles préservées, jamais remplacées
        [Fact]
        public void Render_NamedRangeLoopWithPartiallyOutsideCF_PreservesRealColumns()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            BuildNamedRangeTemplate(sheet);
            var cf = sheet.ConditionalFormatting.AddExpression("E10:H12");
            cf.Formula = "E11>100";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items } }, sheet);

            var rule = Assert.Single(sheet.ConditionalFormatting);
            Assert.Equal(5, rule.Address.Start.Column);
            Assert.Equal(8, rule.Address.End.Column);
            Assert.True(rule.Address.Start.Row <= 10, "CF should start at or before the block start row");
            Assert.True(rule.Address.End.Row >= 16, "CF should cover the final block rows");
        }

        // Raffinement : D1:D1000 est une plage finie réconciliable, distincte du passthrough D:D
        [Fact]
        public void Render_NamedRangeLoopWithFiniteRangeCF_ExtendsUnlikeWholeColumn()
        {
            SetupLicense();
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            BuildNamedRangeTemplate(sheet);
            var cf = sheet.ConditionalFormatting.AddExpression("D1:D1000");
            cf.Formula = "D11>100";

            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            var renderer = new TemplateRenderer(new ExpressionEvaluator());
            renderer.Render(template, new RenderContext { Current = new { Items } }, sheet);

            var rule = Assert.Single(sheet.ConditionalFormatting);
            Assert.Equal(4, rule.Address.Start.Column);
            Assert.Equal(4, rule.Address.End.Column);
            Assert.Equal(1, rule.Address.Start.Row);
            Assert.True(rule.Address.End.Row >= 1001,
                "Finite range must follow the block extension instead of staying untouched like D:D");
        }
    }
}
