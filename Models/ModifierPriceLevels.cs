namespace MSaleWebServer.Models
{
    public class ModifierPriceLevelDto
    {
        public string ModifierTitle { get; set; }
        public string Prefix { get; set; }
        public decimal Price { get; set; }
        public int Level { get; set; }
        public string ModifierCategory { get; set; }
    }
}
