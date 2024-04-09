using System;
using System.Collections.Generic;

namespace ReadExcelFile.Lib.Models
{
    public class WorkSheetData
    {
        public List<String> ColumnHeaders { get; set; }
        public List<Toy> Toys { get; set; }
        public List<Game> Games { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public WorkSheetData()
        {
            ColumnHeaders = new List<String>();
            Games = new List<Game>();
            Toys = new List<Toy>();
        }
    }
}
