namespace ReadExcelFile.Lib.Models
{
    /// <summary>
    /// Store unique game fields
    /// </summary>
    public class Game : Universal
    {
        public bool Sealed { get; set; }
        public string Platform { get; set; }
        public string MediaType { get; set; }
    }
}
