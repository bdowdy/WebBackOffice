namespace MSaleWebServer.Models
{
    public class TenderDto
    {
        public string TenderName { get; set; }

        public bool UsesEMV { get; set; }
        public bool IsGiftEPay { get; set; }
        public bool IsHouseAccount { get; set; }
        public bool IsKiosk { get; set; }
        public bool IsInterCard { get; set; }
        public bool IsPropertyManagement { get; set; }
        public bool IsCash { get; set; }
        public bool OpensDrawer { get; set; }
        public bool ManagerAccess { get; set; }

        public decimal ConvenienceFeeAmt { get; set; }
        public string FeeType { get; set; }
    }
}
