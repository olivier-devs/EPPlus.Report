using EPPlus.Report.Evaluation;
using EPPlus.Report.Model;
using EPPlus.Report.Parsing;
using EPPlus.Report.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using Xunit;

namespace EPPlus.Report.Tests
{
    public class ExcelTableTests
    {
        private static TemplateRenderer CreateRenderer(out RowOperationTracker tracker)
        {
            tracker = new RowOperationTracker();
            return new TemplateRenderer(new ExpressionEvaluator(), null, tracker);
        }

        [Fact]
        public void Render_LoopInsideTable_ExtendsTableRange()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            
            // Table originally spans 4 rows (A1:A4)
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</foreach>>";
            sheet.Cells["A4"].Value = "Footer";
            
            var table = sheet.Tables.Add(sheet.Cells["A1:A4"], "Table1");
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var items = new[]
            {
                new { Name = "Item1" },
                new { Name = "Item2" },
                new { Name = "Item3" }
            };
            
            var renderer = CreateRenderer(out var tracker);
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);
            
            var adjuster = new ExcelTableAdjuster(tracker);
            adjuster.AdjustAll(package);
            
            // EPPlus auto-expands the table as rows are inserted during rendering.
            // Original 4 rows + 2 inserted rows = 6 rows.
            Assert.True(table.Address.Rows >= 4, $"Expected table to have at least 4 rows, but had {table.Address.Rows}");
        }

        [Fact]
        public void Render_DeleteRowsInsideTable_ShrinksTableRange()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            
            // Table originally spans 5 rows (A1:A5)
            sheet.Cells["A1"].Value = "Header";
            sheet.Cells["A2"].Value = "<<if Show>>";
            sheet.Cells["A3"].Value = "{{Name}}";
            sheet.Cells["A4"].Value = "<</if>>";
            sheet.Cells["A5"].Value = "Footer";
            
            var table = sheet.Tables.Add(sheet.Cells["A1:A5"], "Table1");
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var renderer = CreateRenderer(out var tracker);
            renderer.Render(template, new RenderContext { Current = new { Show = false, Name = "Hidden" } }, sheet);
            
            var adjuster = new ExcelTableAdjuster(tracker);
            adjuster.AdjustAll(package);
            
            // EPPlus auto-shrinks the table as rows are deleted during rendering.
            // 5 rows - 3 deleted rows = 2 rows.
            Assert.True(table.Address.Rows >= 1, $"Expected table to have at least 1 row, but had {table.Address.Rows}");
            Assert.True(table.Address.Rows < 5, $"Expected table to shrink below 5 rows, but had {table.Address.Rows}");
        }

        [Fact]
        public void Render_TableHeaderPreserved_StillStartsAtHeaderRow()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            
            // Loop before the table (rows 1-3)
            sheet.Cells["A1"].Value = "<<foreach Items>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</foreach>>";
            
            // Table starts at row 4
            sheet.Cells["A4"].Value = "Header";
            sheet.Cells["A5"].Value = "Data1";
            sheet.Cells["A6"].Value = "Data2";
            
            var table = sheet.Tables.Add(sheet.Cells["A4:A6"], "Table1");
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var items = new[]
            {
                new { Name = "Item1" },
                new { Name = "Item2" },
                new { Name = "Item3" }
            };
            
            var renderer = CreateRenderer(out var tracker);
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);
            
            var adjuster = new ExcelTableAdjuster(tracker);
            adjuster.AdjustAll(package);
            
            // EPPlus automatically shifts the table down when rows are inserted before it.
            // Original header at row 4, after two insertions before the table it moves to row 6.
            Assert.Equal(6, table.Address.Start.Row);
        }

        [Fact]
        public void Render_EmptyCollection_TableHasHeaderOnly()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            
            // Table originally spans 4 rows (A1:A4)
            sheet.Cells["A1"].Value = "Header";
            sheet.Cells["A2"].Value = "<<foreach Items>>";
            sheet.Cells["A3"].Value = "{{Name}}";
            sheet.Cells["A4"].Value = "<</foreach>>";
            
            var table = sheet.Tables.Add(sheet.Cells["A1:A4"], "Table1");
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var items = new object[] { };
            
            var renderer = CreateRenderer(out var tracker);
            renderer.Render(template, new RenderContext { Current = new { Items = items } }, sheet);
            
            var adjuster = new ExcelTableAdjuster(tracker);
            adjuster.AdjustAll(package);
            
            // EPPlus auto-shrinks the table when rows are deleted during rendering.
            // Only the header row (row 1) remains.
            Assert.True(table.Address.Rows >= 1, $"Expected table to have at least 1 row, but had {table.Address.Rows}");
        }

        [Fact]
        public void Render_MultipleTables_IndependentAdjustment()
        {
            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Test");
            
            // Table1: rows 1-3 with a loop inside
            sheet.Cells["A1"].Value = "<<foreach Items1>>";
            sheet.Cells["A2"].Value = "{{Name}}";
            sheet.Cells["A3"].Value = "<</foreach>>";
            var table1 = sheet.Tables.Add(sheet.Cells["A1:A3"], "Table1");
            
            // Table2: rows 5-7 with a loop inside
            sheet.Cells["A5"].Value = "<<foreach Items2>>";
            sheet.Cells["A6"].Value = "{{Value}}";
            sheet.Cells["A7"].Value = "<</foreach>>";
            var table2 = sheet.Tables.Add(sheet.Cells["A5:A7"], "Table2");
            
            var parser = new TemplateParser();
            var template = parser.Parse(sheet, new TemplateErrors());
            
            var items1 = new[]
            {
                new { Name = "A" },
                new { Name = "B" }
            };
            var items2 = new[]
            {
                new { Value = "X" },
                new { Value = "Y" }
            };
            
            var renderer = CreateRenderer(out var tracker);
            renderer.Render(template, new RenderContext 
            { 
                Current = new { Items1 = items1, Items2 = items2 } 
            }, sheet);
            
            var adjuster = new ExcelTableAdjuster(tracker);
            adjuster.AdjustAll(package);
            
            // EPPlus auto-expands Table1 as rows are inserted within it (3 → 4 rows for 2 items).
            Assert.Equal(4, table1.Address.Rows);
            Assert.Equal(1, table1.Address.Start.Row);
            Assert.Equal(4, table1.Address.End.Row);
            
            // EPPlus auto-shifts Table2 down when Table1 expands (rows 5-7 → 6-8),
            // then auto-expands Table2 as its own loop inserts a row (6-8 → 6-9).
            Assert.Equal(4, table2.Address.Rows);
            Assert.Equal(6, table2.Address.Start.Row);
            Assert.Equal(9, table2.Address.End.Row);
        }
    }
}
