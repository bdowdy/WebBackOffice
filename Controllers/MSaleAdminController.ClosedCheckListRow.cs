namespace MSaleWebServer.Controllers
{
    public partial class MSaleAdminController
    {
        public class ClosedCheckListRow
        {
            public int CheckNumber { get; set; }
            public string ServerName { get; set; } = "";
            public string CashierName { get; set; } = "";
            public string OrderDate { get; set; } = "";
            public string OpenTime { get; set; } = "";
            public string CloseTime { get; set; } = "";
            public decimal CheckTotal { get; set; }
            public int CloseOutDay { get; set; }
        }

    }
}
