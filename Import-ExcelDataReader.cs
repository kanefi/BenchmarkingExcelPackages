using ExcelDataReader;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BenchmarkingExcelPackages
{
    class Import_ExcelDataReader
    {
        public DataSet ReadDataFromFile()

        {
            // read excel file
            var filePath = @"C:\Users\FKANE\source\repos\BenchmarkingExcelPackages\lotsofdata.xlsx";

            // create new excel package in a memory stream
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Use the AsDataSet extension method

                    DataSet result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                    var DataReadFromFile = result.Tables;
                }
            }
            //return 
        }

        public void WriteToNewFile()
        {
            var data = ReadDataFromFile();
        }




    }
}
