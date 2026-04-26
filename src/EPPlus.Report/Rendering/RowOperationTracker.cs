using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace EPPlus.Report.Rendering;

/// <summary>
///     Represents a single row insertion or deletion operation performed during template rendering.
/// </summary>
public class RowOperation
{
    /// <summary>
    ///     Gets or sets the worksheet where the operation occurred.
    /// </summary>
    public ExcelWorksheet Worksheet { get; set; }

    /// <summary>
    ///     Gets or sets the row index where the operation started.
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    ///     Gets or sets the number of rows affected.
    /// </summary>
    public int Count { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether the operation was an insertion; <c>false</c> indicates a deletion.
    /// </summary>
    public bool IsInsert { get; set; }
}

/// <summary>
///     Tracks row insertions and deletions performed during template rendering for post-processing adjustments.
/// </summary>
public class RowOperationTracker
{
    private readonly List<RowOperation> _operations = [];

    /// <summary>
    ///     Records the insertion of rows in a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet where rows were inserted.</param>
    /// <param name="row">The row index where insertion started.</param>
    /// <param name="count">The number of rows inserted.</param>
    public void RecordInsert(ExcelWorksheet worksheet, int row, int count)
    {
        _operations.Add(new RowOperation { Worksheet = worksheet, Row = row, Count = count, IsInsert = true });
    }

    /// <summary>
    ///     Records the deletion of rows in a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet where rows were deleted.</param>
    /// <param name="row">The row index where deletion started.</param>
    /// <param name="count">The number of rows deleted.</param>
    public void RecordDelete(ExcelWorksheet worksheet, int row, int count)
    {
        _operations.Add(new RowOperation { Worksheet = worksheet, Row = row, Count = count, IsInsert = false });
    }

    /// <summary>
    ///     Gets all recorded operations for the specified worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to filter by.</param>
    /// <returns>A read-only list of <see cref="RowOperation" /> instances for the worksheet.</returns>
    public IReadOnlyList<RowOperation> GetOperations(ExcelWorksheet worksheet)
    {
        return _operations.Where(o => o.Worksheet == worksheet).ToList();
    }
}