namespace MSaleWebServer.Models
{
    public class DiscountEventAssignmentSaveModel
    {
        public string EventName { get; set; } = "";
        public List<string> DiscountNames { get; set; } = new();
    }
}
