namespace MSaleWebServer.Models
{
    public class DiscountDayOption
    {
        public int Pos { get; set; }
        public string Day { get; set; } = "";
        public bool Selected { get; set; }
    }
}