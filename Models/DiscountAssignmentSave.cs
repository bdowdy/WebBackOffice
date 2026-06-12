namespace MSaleWebServer.Models
{
    public class DiscountAssignmentSaveModel
    {
        public string DiscountName { get; set; }
        public List<DiscountAssignmentItem> Items { get; set; }
    }

    public class DiscountAssignmentItem
    {
        public string ItemName { get; set; }
        public string ItemType { get; set; }
    }
}
