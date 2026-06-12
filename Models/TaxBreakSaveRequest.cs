namespace MSaleWebServer.Models
{
    public class TaxBreakSaveRequest
    {
        public int TaxTable { get; set; }
        public List<TaxBreakRow> Rows { get; set; } = new();
    }

    public class TaxBreakRow
    {
        public int RowNumber { get; set; }
        public decimal StartAmount { get; set; }
        public decimal EndAmount { get; set; }
        public decimal TaxAmount { get; set; }
    }
}
