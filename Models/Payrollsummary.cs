namespace MSaleWebServer.Models
{
    public class PayrollSummaryRow
    {
        public string EmployeeName { get; set; } = "";
        public string DepartmentName { get; set; } = "";
        public DateTime ShiftDate { get; set; }
        public decimal RateOfPay { get; set; }
        public decimal RegHours { get; set; }
        public decimal OTHours { get; set; }
        public decimal TotalHours { get; set; }
        public decimal Wages { get; set; }
        public decimal DeclaredTips { get; set; }
        public decimal CreditCardTips { get; set; }
        public decimal TipsPaidOut { get; set; }
        public decimal TipShare { get; set; }
        public decimal Sales { get; set; }
    }
}
