namespace MSaleWebServer.Models
{
    public class ButtonLayoutSaveRequest
    {
        public string Category { get; set; } = "";
        public List<ButtonLayoutDto> Buttons { get; set; } = new();
    }

    public class ButtonLayoutDto
    {
        public string MenuItem { get; set; } = "";
        public string Category { get; set; } = "";
        public string ForeCol { get; set; } = "16777215";
        public string BackCol { get; set; } = "4210752";
        public string Icon { get; set; } = "";
        public int? Width { get; set; }
        public int? Height { get; set; }
        public int? FontSize { get; set; }
        public int? Left { get; set; }
        public int? Top { get; set; }
        public int? ButtonAssigned { get; set; }
    }
}
