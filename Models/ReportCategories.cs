namespace MSaleWebServer.Models
{
    public class ReportCategoryDto
    {
        public string ReportName { get; set; } = "";

        public int Flag { get; set; }
        public int RecordNumber { get; set; }
        public int DepartmentNo { get; set; }

        public string Remote1 { get; set; } = "";
        public string Remote2 { get; set; } = "";
        public string Remote3 { get; set; } = "";
        public string Remote4 { get; set; } = "";
        public string Remote5 { get; set; } = "";
        public string Remote6 { get; set; } = "";

        public string MasterName { get; set; } = "None";

        public string TimedEvent1 { get; set; } = "None";
        public string TimedEvent2 { get; set; } = "None";
        public string TimedEvent3 { get; set; } = "None";
        public string TimedEvent4 { get; set; } = "None";
        public string TimedEvent5 { get; set; } = "None";

       
    }
}
