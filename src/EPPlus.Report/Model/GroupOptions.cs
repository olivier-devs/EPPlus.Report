namespace EPPlus.Report.Model
{
    public class GroupOptions
    {
        public bool Collapse { get; set; }
        public MergeMode MergeLabels { get; set; } = MergeMode.None;
        public int PlaceToColumn { get; set; }
        public bool WithHeader { get; set; }
        public bool DisableSubtotals { get; set; }
        public bool DisableOutline { get; set; }
        public bool PageBreaks { get; set; }
        public string TotalLabel { get; set; } = "Total";
        public string GrandLabel { get; set; } = "Grand";
        public bool SummaryAbove { get; set; }
        public bool DisableGrandTotal { get; set; }
        public bool Descending { get; set; }
    }
}
