using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.Report.Rendering;

/// <summary>
///     Adjusts Excel table ranges after row insertions and deletions performed during rendering.
/// </summary>
public class ExcelTableAdjuster
{
    private readonly RowOperationTracker _tracker;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelTableAdjuster" /> class.
    /// </summary>
    /// <param name="tracker">The tracker containing row operation history.</param>
    public ExcelTableAdjuster(RowOperationTracker tracker)
    {
        _tracker = tracker;
    }

    /// <summary>
    ///     Adjusts all tables in all worksheets of the package based on tracked row operations.
    /// </summary>
    /// <param name="package">The Excel package to adjust.</param>
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

    private static void AdjustTable(ExcelTable table, ExcelWorksheet worksheet)
    {
        // EPPlus 7+ automatically adjusts table ranges when rows are inserted or deleted.
        // No manual adjustment is needed; doing so would cause double-counting.
    }
}