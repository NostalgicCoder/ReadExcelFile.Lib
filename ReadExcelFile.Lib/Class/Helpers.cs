using System;

namespace ReadExcelFile.Lib.Class
{
    public class Helpers
    {
        /// <summary>
        /// Convert a spreadsheet 'Damaged' field result to more database friendly result
        /// </summary>
        /// <param name="cellVal"></param>
        /// <returns></returns>
        public string ConvertStringToDamaged(string cellVal)
        {
            switch (cellVal.ToString().ToLower())
            {
                case "yes":
                    return "Yes";
                    break;
                case "no":
                    return "No";
                    break;
                case "?":
                    return "Unknown";
                    break;
                default:
                    return "No";
                    break;
            }
        }

        /// <summary>
        /// Convert a spreadsheet 'Complete' field result to more database friendly result
        /// </summary>
        /// <param name="cellVal"></param>
        /// <returns></returns>
        public string ConvertStringToComplete(string cellVal)
        {
            switch (cellVal.ToString().ToLower())
            {
                case "yes":
                    return "Yes";
                    break;
                case "no":
                    return "No";
                    break;
                case "?":
                    return "Unknown";
                    break;
                default:
                    return "Unknown";
                    break;
            }
        }

        /// <summary>
        /// Convert a spreadsheet 'MediaType' field result to more database friendly result
        /// </summary>
        /// <param name="cellVal"></param>
        /// <returns></returns>
        public string ConvertStringToMediaType(string cellVal)
        {
            switch (cellVal.ToString().ToLower())
            {
                case "3.5\"":
                    return "3.5 Disk";
                    break;
                case "5.25\"":
                    return "5.25 Disk";
                    break;
                default:
                    return cellVal;
                    break;
            }
        }

        /// <summary>
        /// Convert abbreviated string representation of a month to a integer, also cater for different representations of a month result
        /// </summary>
        /// <param name="cellVal"></param>
        /// <returns></returns>
        public Int32 ConvertStringMonthToInt32(string cellVal)
        {
            switch (cellVal.ToLower())
            {
                case "jan":
                    return 1;
                    break;
                case "feb":
                    return 2;
                    break;
                case "mar":
                    return 3;
                    break;
                case "apr":
                    return 4;
                    break;
                case "may":
                    return 5;
                    break;
                case "jun":
                    return 6;
                    break;
                case "jul":
                    return 7;
                case "july":
                    return 7;
                    break;
                case "aug":
                    return 8;
                    break;
                case "sep":
                    return 9;
                case "sept":
                    return 9;
                    break;
                case "oct":
                    return 10;
                    break;
                case "nov":
                    return 11;
                    break;
                case "dec":
                    return 12;
                    break;
                default:
                    return 1;
                    break;
            }
        }
    }
}
