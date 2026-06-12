namespace MSaleWebServer.Models
{
    public class TimeMessageDto
    {
        public string DeliveryDate { get; set; } = "";
        public string Message { get; set; } = "";
        public string MessageFrom { get; set; } = "MSaleWebServer";
        public List<string> Names { get; set; } = new();
    }
}
