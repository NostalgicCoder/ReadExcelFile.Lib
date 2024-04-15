namespace ReadExcelFile.Lib.Models
{
    /// <summary>
    /// Store universal and unique toy fields
    /// </summary>
    public class Toy : Universal
    {
        // Universal Toy Fields
        public string Damaged { get; set; }
        public string DamagedAccessory { get; set; }

        // Unique Toy Fields
        public string Stands { get; set; }
        public string Colour { get; set; }
        public string ToyLine { get; set; }
        public bool Carded { get; set; }
        public bool Boxed { get; set; }
        public bool Discoloured { get; set; }
    }
}