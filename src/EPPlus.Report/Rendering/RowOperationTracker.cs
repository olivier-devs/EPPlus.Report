using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace EPPlus.Report.Rendering
{
    public class RowOperation
    {
        public ExcelWorksheet Worksheet { get; set; }
        public int Row { get; set; }
        public int Count { get; set; }
        public bool IsInsert { get; set; }
    }

    public class RowOperationTracker
    {
        private readonly List<RowOperation> _operations = new List<RowOperation>();

        public void RecordInsert(ExcelWorksheet worksheet, int row, int count)
        {
            _operations.Add(new RowOperation { Worksheet = worksheet, Row = row, Count = count, IsInsert = true });
        }

        public void RecordDelete(ExcelWorksheet worksheet, int row, int count)
        {
            _operations.Add(new RowOperation { Worksheet = worksheet, Row = row, Count = count, IsInsert = false });
        }

        public IReadOnlyList<RowOperation> GetOperations(ExcelWorksheet worksheet)
        {
            return _operations.Where(o => o.Worksheet == worksheet).ToList();
        }
    }
}
