namespace MSaleWebServer.Models
{
    public class MealPeriodModel
    {
        public string PeriodName { get; set; }

        public string PeriodTime { get; set; }
    }

    public class CategoryMealPeriodModel
    {
        public string CategoryName { get; set; }

        public string MealPeriod { get; set; }
    }
}
