namespace EPPlus.Report
{
    public class TemplateError
    {
        public string Message { get; set; }
        public string CellAddress { get; set; }
        public string WorksheetName { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public string Expression { get; set; }
        public ErrorType Type { get; set; }

        public string Location => $"{WorksheetName}!{CellAddress}";
    }
}
