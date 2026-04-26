using OfficeOpenXml;
using System.Linq;

namespace EPPlus.Report.Rendering
{
    public class ExcelTableAdjuster
    {
        private readonly RowOperationTracker _tracker;

        public ExcelTableAdjuster(RowOperationTracker tracker)
        {
            _tracker = tracker;
        }

        public void AdjustAll(ExcelPackage package)
        {
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                foreach (var table in worksheet.Tables.ToList())
                {
                    AdjustTable(table, worksheet);
                }
            }
        }

        private void AdjustTable(OfficeOpenXml.Table.ExcelTable table, ExcelWorksheet worksheet)
        {
            // EPPlus 7+ automatically adjusts table ranges when rows are inserted or deleted.
            // No manual adjustment is needed; doing so would cause double-counting.
        }
    }
}
