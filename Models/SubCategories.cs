namespace MSaleWebServer.Models
{
    public class SubCategoryDto
    {
        public string SubName { get; set; } = "";
        public string CategoryName { get; set; } = "";
        public string TimedEvent1 { get; set; } = "None";
        public string TimedEvent2 { get; set; } = "None";
        public string TimedEvent3 { get; set; } = "None";
        public string TimedEvent4 { get; set; } = "None";
        public string TimedEvent5 { get; set; } = "None";
    }

    public class RenameSubCategoryDto
    {
        public string OldName { get; set; } = "";
        public string NewName { get; set; } = "";
    }
}
