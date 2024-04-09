using OfficeOpenXml;
using ReadExcelFile.Lib.Class;
using ReadExcelFile.Lib.Enum;
using ReadExcelFile.Lib.Models;
using System;
using System.Collections.Generic;
using System.IO;

// https://coderwall.com/p/app3ya/read-excel-file-in-c
// https://stackoverflow.com/questions/44916744/do-i-need-to-have-office-installed-to-use-microsoft-office-interop-excel-dll

namespace ReadExcelFile.Lib
{
    public class ProcessExcelFile
    {
        private Helpers _helpers = new Helpers();
        private WorkSheetData _workSheetData = new WorkSheetData();

        private List<Game> _gameColl = new List<Game>();
        private List<Toy> _toyColl = new List<Toy>();

        private FileInfo _existingFile = new FileInfo(@"D:\My Files\Documents\Toys Prices.xlsx");

        /// <summary>
        /// Open the spreadsheet, process each sheet one at a time and output the sheets results to a individual '.SQL' script file.
        /// </summary>
        public void ReadExcelFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(_existingFile))
            {
                ProcessGame(package, (int)Sheet.Games);
                WriteResultsToASqlScript(_workSheetData, Sheet.Games);

                ProcessToy(package, (int)Sheet.Motu, Sheet.Motu);
                WriteResultsToASqlScript(_workSheetData, Sheet.Motu);

                ProcessToy(package, (int)Sheet.MotuOrigins, Sheet.MotuOrigins);
                WriteResultsToASqlScript(_workSheetData, Sheet.MotuOrigins);

                ProcessToy(package, (int)Sheet.Tmnt, Sheet.Tmnt);
                WriteResultsToASqlScript(_workSheetData, Sheet.Tmnt);

                ProcessToy(package, (int)Sheet.ThunderCats, Sheet.ThunderCats);
                WriteResultsToASqlScript(_workSheetData, Sheet.ThunderCats);

                ProcessToy(package, (int)Sheet.Mask, Sheet.Mask);
                WriteResultsToASqlScript(_workSheetData, Sheet.Mask);

                ProcessToy(package, (int)Sheet.Mimp, Sheet.Mimp);
                WriteResultsToASqlScript(_workSheetData, Sheet.Mimp);

                ProcessToy(package, (int)Sheet.SuperPowers, Sheet.SuperPowers);
                WriteResultsToASqlScript(_workSheetData, Sheet.SuperPowers);

                ProcessToy(package, (int)Sheet.Simpsons, Sheet.Simpsons);
                WriteResultsToASqlScript(_workSheetData, Sheet.Simpsons);

                ProcessToy(package, (int)Sheet.Misc, Sheet.Misc);
                WriteResultsToASqlScript(_workSheetData, Sheet.Misc);
            }

            Console.WriteLine("Finished Data Export!");
        }

        /// <summary>
        /// Read a single sheet in the spreadsheet, refine and map the cell data to a model and store as a collection of results.
        /// </summary>
        /// <param name="package"></param>
        /// <param name="workSheet"></param>
        /// <param name="sheet"></param>
        private void ProcessToy(ExcelPackage package, int workSheet, Sheet sheet)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[workSheet];
            List<String> columnHeaders = new List<string>();

            // Clear any previous results from an earlier run
            _workSheetData.Toys.Clear();

            _workSheetData.ColumnHeaders = columnHeaders;

            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                // Handle if there is no value found in the cell
                if (worksheet.Cells[1, col].Value != null)
                {
                    columnHeaders.Add(worksheet.Cells[1, col].Value.ToString().Trim());
                }
            }

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                Toy toy = new Toy();

                if (worksheet.Cells[row, 1].Value == null)
                {
                    // Worksheet has run out of results, so end the loop
                    break;
                }

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    // Handle if there is no value found in the cell
                    if (worksheet.Cells[row, col].Value != null)
                    {
                        switch (sheet)
                        {
                            case Sheet.Motu:
                                MapMotuCellValues(toy, worksheet.Cells[row, col].Value.ToString(), col);
                                break;
                            case Sheet.MotuOrigins:
                                MapMotuOriginsCellValues(toy, worksheet.Cells[row, col].Value.ToString(), col);
                                break;
                            case Sheet.Tmnt:
                                MapTmntCellValues(toy, worksheet.Cells[row, col].Value.ToString(), col);
                                break;
                            case Sheet.ThunderCats :
                                MapThunderCatsCellValues(toy, worksheet.Cells[row, col].Value.ToString(), col);
                                break;
                            case Sheet.Mask:
                                MapTmntCellValues(toy, worksheet.Cells[row, col].Value.ToString(), col);
                                break;
                            case Sheet.Mimp:
                                MapMimpCellValues(toy, worksheet.Cells[row, col].Value.ToString(), col);
                                break;
                            case Sheet.SuperPowers:
                                MapSuperPowersCellValues(toy, worksheet.Cells[row, col].Value.ToString(), col);
                                break;
                            case Sheet.Simpsons:
                                MapSuperPowersCellValues(toy, worksheet.Cells[row, col].Value.ToString(), col);
                                break;
                            case Sheet.Misc:
                                MapMiscCellValues(toy, worksheet.Cells[row, col].Value.ToString(), col);
                                break;
                        }
                    }
                    else
                    {
                        switch (sheet)
                        {
                            case Sheet.Motu:
                                MapMotuCellValues(toy, string.Empty, col);
                                break;
                            case Sheet.MotuOrigins:
                                MapMotuOriginsCellValues(toy, string.Empty, col);
                                break;
                            case Sheet.Tmnt:
                                MapTmntCellValues(toy, string.Empty, col);
                                break;
                            case Sheet.ThunderCats:
                                MapThunderCatsCellValues(toy, string.Empty, col);
                                break;
                            case Sheet.Mask:
                                MapTmntCellValues(toy, string.Empty, col);
                                break;
                            case Sheet.Mimp:
                                MapMimpCellValues(toy, string.Empty, col);
                                break;
                            case Sheet.SuperPowers:
                                MapSuperPowersCellValues(toy, string.Empty, col);
                                break;
                            case Sheet.Simpsons:
                                MapSuperPowersCellValues(toy, string.Empty, col);
                                break;
                            case Sheet.Misc:
                                MapMiscCellValues(toy, string.Empty, col);
                                break;
                        }
                    }
                }

                _toyColl.Add(toy);
            }

            foreach (Toy x in _toyColl)
            {
                x.SaleDate = new DateTime(x.Year, x.Month, 1);
            }

            _workSheetData.Toys = _toyColl;
        }

        private void ProcessGame(ExcelPackage package, int workSheet)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[workSheet];
            List<String> columnHeaders = new List<string>();

            _workSheetData.ColumnHeaders = columnHeaders;

            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                // Handle if there is no value found in the cell
                if (worksheet.Cells[1, col].Value != null)
                {
                    columnHeaders.Add(worksheet.Cells[1, col].Value.ToString().Trim());
                }
            }

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                Game game = new Game();

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    // Handle if there is no value found in the cell
                    if (worksheet.Cells[row, col].Value != null)
                    {
                        MapGameCellValues(game, worksheet.Cells[row, col].Value.ToString(), col);
                    }
                    else
                    {
                        MapGameCellValues(game, string.Empty, col);
                    }
                }

                _gameColl.Add(game);
            }

            foreach (Game x in _gameColl)
            {
                x.SaleDate = new DateTime(x.Year, x.Month, 1);
            }

            _workSheetData.Games = _gameColl;
        }

        #region Map Cells To Model

        private void MapGameCellValues(Game game, string cellVal, int col)
        {
            try
            {
                switch (col)
                {
                    case 1:
                        game.Name = cellVal.ToString().Replace("'", "''");
                        break;
                    case 2:
                        game.Condition = cellVal;
                        break;
                    case 3:
                        game.Sealed = (cellVal.ToLower() == "yes") ? true : false;
                        break;
                    case 4:
                        game.Platform = cellVal;
                        break;
                    case 5:
                        game.MediaType = _helpers.ConvertStringMediaType(cellVal);
                        break;
                    case 6:
                        game.Complete = _helpers.ConvertStringComplete(cellVal);
                        break;
                    case 7:
                        game.Price = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 8:
                        game.Postage = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 10:
                        game.Month = _helpers.ConvertStringMonthToInt32(cellVal);
                        break;
                    case 11:
                        game.Year = string.IsNullOrEmpty(cellVal) ? 0001 : Convert.ToInt32(cellVal);
                        break;
                    case 12:
                        game.Description = cellVal.Replace("'", "''");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void MapMotuCellValues(Toy toy, string cellVal, int col)
        {
            try
            {
                switch (col)
                {
                    case 1:
                        toy.Name = cellVal.ToString().Replace("'", "''");
                        break;
                    case 2:
                        toy.Condition = cellVal;
                        break;
                    case 3:
                        toy.Damaged = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 4:
                        toy.DamagedAccessory = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 5:
                        toy.Stands = _helpers.ConvertStringComplete(cellVal);
                        break;
                    case 6:
                        toy.Complete = _helpers.ConvertStringComplete(cellVal);
                        break;
                    case 7:
                        toy.Price = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 8:
                        toy.Postage = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 10:
                        toy.Month = _helpers.ConvertStringMonthToInt32(cellVal);
                        break;
                    case 11:
                        toy.Year = string.IsNullOrEmpty(cellVal) ? 0001 : Convert.ToInt32(cellVal);
                        break;
                    case 12:
                        toy.Description = cellVal.ToString().Replace("'", "''"); ;
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void MapMotuOriginsCellValues(Toy toy, string cellVal, int col)
        {
            try
            {
                switch (col)
                {
                    case 1:
                        toy.Name = cellVal.ToString().Replace("'", "''");
                        break;
                    case 2:
                        toy.Carded = (cellVal.ToLower() == "yes") ? true : false;
                        break;
                    case 3:
                        toy.Condition = cellVal;
                        break;
                    case 4:
                        toy.Complete = _helpers.ConvertStringComplete(cellVal);
                        break;
                    case 5:
                        toy.Price = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 6:
                        toy.Postage = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 8:
                        toy.Month = _helpers.ConvertStringMonthToInt32(cellVal);
                        break;
                    case 9:
                        toy.Year = string.IsNullOrEmpty(cellVal) ? 0001 : Convert.ToInt32(cellVal);
                        break;
                    case 10:
                        toy.Description = cellVal.ToString().Replace("'", "''"); ;
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void MapTmntCellValues(Toy toy, string cellVal, int col)
        {
            try
            {
                switch (col)
                {
                    case 1:
                        toy.Name = cellVal.ToString().Replace("'", "''");
                        break;
                    case 2:
                        toy.Condition = cellVal;
                        break;
                    case 3:
                        toy.Damaged = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 4:
                        toy.DamagedAccessory = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 5:
                        toy.Complete = _helpers.ConvertStringComplete(cellVal);
                        break;
                    case 6:
                        toy.Boxed = (cellVal.ToLower() == "yes") ? true : false;
                        break;
                    case 7:
                        toy.Price = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 8:
                        toy.Postage = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 10:
                        toy.Month = _helpers.ConvertStringMonthToInt32(cellVal);
                        break;
                    case 11:
                        toy.Year = string.IsNullOrEmpty(cellVal) ? 0001 : Convert.ToInt32(cellVal);
                        break;
                    case 12:
                        toy.Description = cellVal.ToString().Replace("'", "''"); ;
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void MapThunderCatsCellValues(Toy toy, string cellVal, int col)
        {
            try
            {
                switch (col)
                {
                    case 1:
                        toy.Name = cellVal.ToString().Replace("'", "''");
                        break;
                    case 2:
                        toy.Condition = cellVal;
                        break;
                    case 3:
                        toy.Damaged = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 4:
                        toy.DamagedAccessory = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 5:
                        toy.Complete = _helpers.ConvertStringComplete(cellVal);
                        break;
                    case 6:
                        toy.Price = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 7:
                        toy.Postage = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 9:
                        toy.Month = _helpers.ConvertStringMonthToInt32(cellVal);
                        break;
                    case 10:
                        toy.Year = string.IsNullOrEmpty(cellVal) ? 0001 : Convert.ToInt32(cellVal);
                        break;
                    case 11:
                        toy.Description = cellVal.ToString().Replace("'", "''"); ;
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void MapMimpCellValues(Toy toy, string cellVal, int col)
        {
            try
            {
                switch (col)
                {
                    case 1:
                        toy.Name = cellVal.ToString().Replace("'", "''");
                        break;
                    case 2:
                        toy.Colour = cellVal;
                        break;
                    case 3:
                        toy.Condition = cellVal;
                        break;
                    case 4:
                        toy.Discoloured = (cellVal.ToLower() == "yes") ? true : false;
                        break;
                    case 5:
                        toy.Damaged = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 6:
                        toy.Price = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 7:
                        toy.Postage = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 9:
                        toy.Month = _helpers.ConvertStringMonthToInt32(cellVal);
                        break;
                    case 10:
                        toy.Year = string.IsNullOrEmpty(cellVal) ? 0001 : Convert.ToInt32(cellVal);
                        break;
                    case 11:
                        toy.Description = cellVal.ToString().Replace("'", "''"); ;
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void MapSuperPowersCellValues(Toy toy, string cellVal, int col)
        {
            try
            {
                switch (col)
                {
                    case 1:
                        toy.Name = cellVal.ToString().Replace("'", "''");
                        break;
                    case 2:
                        toy.Condition = cellVal;
                        break;
                    case 3:
                        toy.Carded = (cellVal.ToLower() == "yes") ? true : false;
                        break;
                    case 4:
                        toy.Damaged = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 5:
                        toy.DamagedAccessory = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 6:
                        toy.Complete = _helpers.ConvertStringComplete(cellVal);
                        break;
                    case 7:
                        toy.Price = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 8:
                        toy.Postage = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 10:
                        toy.Month = _helpers.ConvertStringMonthToInt32(cellVal);
                        break;
                    case 11:
                        toy.Year = string.IsNullOrEmpty(cellVal) ? 0001 : Convert.ToInt32(cellVal);
                        break;
                    case 12:
                        toy.Description = cellVal.ToString().Replace("'", "''"); ;
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void MapMiscCellValues(Toy toy, string cellVal, int col)
        {
            try
            {
                switch (col)
                {
                    case 1:
                        string val = cellVal.ToString().Replace("'", "''");

                        val = val.Replace(" ", "");
                        val = val.Replace("-", "_");

                        toy.ToyLine = val;
                        break;
                    case 2:
                        toy.Name = cellVal.ToString().Replace("'", "''");
                        break;
                    case 3:
                        toy.Condition = cellVal;
                        break;
                    case 4:
                        toy.Damaged = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 5:
                        toy.DamagedAccessory = _helpers.ConvertStringDamaged(cellVal);
                        break;
                    case 6:
                        toy.Complete = _helpers.ConvertStringComplete(cellVal);
                        break;
                    case 7:
                        toy.Price = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 8:
                        toy.Postage = (cellVal != string.Empty) ? decimal.Parse(cellVal) : 0.00M;
                        break;
                    case 10:
                        toy.Month = _helpers.ConvertStringMonthToInt32(cellVal);
                        break;
                    case 11:
                        toy.Year = string.IsNullOrEmpty(cellVal) ? 0001 : Convert.ToInt32(cellVal);
                        break;
                    case 12:
                        toy.Description = cellVal.ToString().Replace("'", "''"); ;
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        #endregion

        /// <summary>
        /// Export worksheet data to a '.SQL' file so it can be imported into SQL Express
        /// - The purpose of this is to overcome the limitations of SQL Express and allow me to import data in from a spreadsheet (A feature that's only available on the full price SSMS or on a ETL tool and not on free tooling).
        /// - I searched online and could not find a way to get SQL Express to import data from a spreadsheet file.
        /// </summary>
        /// <param name="workSheetData"></param>
        /// <param name="sheet"></param>
        private void WriteResultsToASqlScript(WorkSheetData workSheetData, Sheet sheet)
        {
            try
            {
                string outputSqlScriptFile = string.Format(@"D:\Downloads\ImportDataScript - {0}.sql", sheet);

                if (File.Exists(outputSqlScriptFile))
                {
                    try
                    {
                        File.Delete(outputSqlScriptFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                }

                using (StreamWriter writer = new StreamWriter(outputSqlScriptFile))
                {
                    switch (sheet)
                    {
                        case Sheet.Motu:
                            foreach (var x in workSheetData.Toys)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Toys (Description, Colour, Damaged, DamagedAccessory, Discoloured, Carded, Boxed, Stands, ToyLine, Name, Condition, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')",
                                    x.Description, null, x.Damaged, x.DamagedAccessory, false, false, false, x.Stands, Sheet.Motu, x.Name, x.Condition, x.Complete, x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                        case Sheet.MotuOrigins:
                            foreach (var x in workSheetData.Toys)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Toys (Description, Colour, Damaged, DamagedAccessory, Discoloured, Carded, Boxed, Stands, ToyLine, Name, Condition, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')",
                                    x.Description, null, false, false, false, x.Carded, false, true, Sheet.MotuOrigins, x.Name, x.Condition, x.Complete, x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                        case Sheet.Tmnt:
                            foreach (var x in workSheetData.Toys)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Toys (Description, Colour, Damaged, DamagedAccessory, Discoloured, Carded, Boxed, Stands, ToyLine, Name, Condition, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')",
                                    x.Description, null, x.Damaged, x.DamagedAccessory, false, false, x.Boxed, true, Sheet.Tmnt, x.Name, x.Condition, x.Complete, x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                        case Sheet.ThunderCats:
                            foreach (var x in workSheetData.Toys)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Toys (Description, Colour, Damaged, DamagedAccessory, Discoloured, Carded, Boxed, Stands, ToyLine, Name, Condition, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')",
                                    x.Description, null, x.Damaged, x.DamagedAccessory, false, false, false, true, Sheet.ThunderCats, x.Name, x.Condition, x.Complete, x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                        case Sheet.Mask:
                            foreach (var x in workSheetData.Toys)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Toys (Description, Colour, Damaged, DamagedAccessory, Discoloured, Carded, Boxed, Stands, ToyLine, Name, Condition, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')",
                                    x.Description, null, x.Damaged, x.DamagedAccessory, false, false, x.Boxed, true, Sheet.Mask, x.Name, x.Condition, x.Complete, x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                        case Sheet.Mimp:
                            foreach (var x in workSheetData.Toys)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Toys (Description, Colour, Damaged, DamagedAccessory, Discoloured, Carded, Boxed, Stands, ToyLine, Name, Condition, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')",
                                    x.Description, x.Colour, x.Damaged, false, x.Discoloured, false, false, true, Sheet.Mimp, x.Name, x.Condition, "Yes", x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                        case Sheet.SuperPowers:
                            foreach (var x in workSheetData.Toys)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Toys (Description, Colour, Damaged, DamagedAccessory, Discoloured, Carded, Boxed, Stands, ToyLine, Name, Condition, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')",
                                    x.Description, null, x.Damaged, x.DamagedAccessory, false, x.Carded, false, true, Sheet.SuperPowers, x.Name, x.Condition, x.Complete, x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                        case Sheet.Simpsons:
                            foreach (var x in workSheetData.Toys)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Toys (Description, Colour, Damaged, DamagedAccessory, Discoloured, Carded, Boxed, Stands, ToyLine, Name, Condition, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')",
                                    x.Description, null, x.Damaged, x.DamagedAccessory, false, x.Carded, false, true, Sheet.Simpsons, x.Name, x.Condition, x.Complete, x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                        case Sheet.Misc:
                            foreach (var x in workSheetData.Toys)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Toys (Description, Colour, Damaged, DamagedAccessory, Discoloured, Carded, Boxed, Stands, ToyLine, Name, Condition, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')",
                                    x.Description, null, x.Damaged, x.DamagedAccessory, false, false, false, true, x.ToyLine, x.Name, x.Condition, x.Complete, x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                        case Sheet.Games:
                            foreach (var x in workSheetData.Games)
                            {
                                writer.WriteLine(string.Format("INSERT INTO Games (Description, Name, Condition, Sealed, Platform, MediaType, Complete, Price, Postage, SaleDate) " +
                                    "VALUES ('{0}','{1}','{2}',{3},'{4}','{5}','{6}','{7}','{8}','{9}')", x.Description, x.Name, x.Condition, Convert.ToInt32(x.Sealed), x.Platform, x.MediaType, x.Complete, x.Price, x.Postage, x.SaleDate.ToString("yyyy-MM-dd HH:mm:ss.fff")));

                                writer.WriteLine("");
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}