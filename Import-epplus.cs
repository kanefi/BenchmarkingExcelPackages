using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BenchmarkingExcelPackages
{
    public class ImportEPPlus
    {
        [Benchmark] // 1.1
        public List<string> ReadDataFromFile()
        {
            List<string> excelData = new List<string>();

            // read excel file
            var excelFile = File.ReadAllBytes(@"C:\Users\FKANE\source\repos\BenchmarkingExcelPackages\lotsofdata.xlsx");

            // create new excel package in a memory stream
            using (MemoryStream stream = new MemoryStream(excelFile))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    // loop all worksheets
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        // 1.6 loop all rows starting from row 2
                        for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)

                            // 1.7 loop all columns in a row
                            for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                            {
                                // add the data to the original list
                                if (worksheet.Cells[i, j].Value != null)
                                {
                                    excelData.Add(worksheet.Cells[i, j].Value.ToString());
                                }
                            }
                        }
                    }
                }
            return excelData;
        }      

        [Benchmark] // 1.2
        public void WriteToNewFile()
        {
            // gather data
            var data = ReadDataFromFile();

            // create a new ExcelPackage
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // create a WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                // 1.3 create a 2nd WorkSheet
                ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets.Add("Sheet 2");

                // add all the data to the excel sheet, starting at cell A1
                worksheet.Cells["A1"].LoadFromCollection(data);

                // add the last row of data to the second worksheet
                worksheet2.Cells["A1"].LoadFromCollection(data[data.Count -1]);

                // 1.4, 1.5 get a range of cells
                var rangeOfCells = worksheet.Cells[2, 2, worksheet.Dimension.End.Row, 2];

                // 1.8 style a cell with color
                rangeOfCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rangeOfCells.Style.Fill.BackgroundColor.SetColor(Color.Red);

                // 1.9 add value to a specific cell 
                var specificCell = worksheet.Cells["F6"];
                specificCell.Value = "Success";

                // 2 style a cell using wingdings font
                rangeOfCells.Style.Font.Name = "WingDings";
                rangeOfCells.Value = "ü";

                // 2.1, 2.4 style text using wrap & alignment
                specificCell.Style.WrapText = true;
                specificCell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                specificCell.Style.TextRotation = 90; //degrees

                // 2.3 bold font
                specificCell.Style.Font.Bold = true;

                // 2.4 merge cells
                specificCell = worksheet.Cells["F6:J6"];
                specificCell.Merge = true;

                // 2.2 style cell border types
                specificCell.Style.Border.Top.Style = ExcelBorderStyle.Double;
                specificCell.Style.Border.Right.Style = ExcelBorderStyle.Double;
                specificCell.Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                specificCell.Style.Border.Left.Style = ExcelBorderStyle.Double;

                // save the newly created file.
                FileInfo fi = new FileInfo(@"C:\Users\FKANE\source\repos\BenchmarkingExcelPackages\CreatedFileFK.xlsx");
                excelPackage.SaveAs(fi);
            }
        }
    }
}