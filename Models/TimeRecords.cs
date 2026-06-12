namespace MSaleWebServer.Models
{
    public class TimeRecordDto
    {
        public string RecordId { get; set; } = "";
        public string EmployeeName { get; set; } = "";
        public string DepartmentName { get; set; } = "";
        public string ShiftDate { get; set; } = "";
        public string TimeIn { get; set; } = "";
        public string BreakOut { get; set; } = "";
        public string BreakIn { get; set; } = "";
        public string TimeOut { get; set; } = "";
        public string PunchOutDate { get; set; } = "";
        public decimal TipAmount { get; set; }
        public decimal RateOfPay { get; set; }
        public decimal Sales { get; set; }
        public decimal CreditCardTips { get; set; }
        public decimal TipsPaidOut { get; set; }
        public decimal TipShare { get; set; }
        public int Shift { get; set; }
    }
}
