namespace MSaleWebServer.Models
{
    public class CouponAssignmentSaveModel
    {
        public string CouponName { get; set; } = "";
        public List<string> LinkToNames { get; set; } = new();
    }

}
