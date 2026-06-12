namespace MSaleWebServer.Models
{
    public class TimedPricingDto
    {
        public string EventName { get; set; } = "";
        public string StartTime { get; set; }
        public string EndTime { get; set; }

        public bool Monday { get; set; }
        public bool Tuesday { get; set; }
        public bool Wednesday { get; set; }
        public bool Thursday { get; set; }
        public bool Friday { get; set; }
        public bool Saturday { get; set; }
        public bool Sunday { get; set; }

        public string StartDate { get; set; }
        public string EndDate { get; set; }

        public decimal PriceOff { get; set; }
    }
}
