using System;

namespace ReadExcelFile.Lib.Models
{
    /// <summary>
    /// Store universal product fields
    /// </summary>
    public class Universal
    {
        public string Name { get; set; }
        public string Condition { get; set; }
        public string Complete { get; set; }
        public decimal Price { get; set; }
        public decimal Postage { get; set; }
        public Int32 Month { get; set; }
        public Int32 Year { get; set; }
        public DateTime SaleDate { get; set; }
        public string Description { get; set; }
    }
}