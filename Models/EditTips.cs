namespace MSaleWebServer.Models
{
    public class TipEditRow
    {
        public string RecordId { get; set; } = "";
        public string EmployeeName { get; set; } = "";
        public string DepartmentName { get; set; } = "";
        public string ShiftDate { get; set; } = "";
        public string TimeIn { get; set; } = "";
        public string TimeOut { get; set; } = "";
        public decimal CashTips { get; set; }
        public decimal CreditCardTips { get; set; }
        public decimal TipsPaidOut { get; set; }
        public decimal TipShare { get; set; }
        public decimal Gratuity { get; set; }
        public int Shift { get; set; }
    }
}
