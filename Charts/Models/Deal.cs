
namespace Charts.Models
{
    public class Deal
    {
        public string Region { get; set; }
        public decimal Value { get; set; }
    }

    public class ChartData
    {
        public string[] Lables; 
        public string[][] Table { get; set; }

    }
}
