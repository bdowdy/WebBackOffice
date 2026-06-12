namespace MSaleWebServer.Models
{
    public class RevenueAssignmentModel
    {
        public string RevenueName { get; set; }

        public string LinkType { get; set; }

        public List<string> LinkToNames { get; set; }
    }
    public class AddRevenueModel
    {
        public string Name { get; set; }
    }
}
