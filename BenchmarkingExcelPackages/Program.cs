using BenchmarkDotNet.Running;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BenchmarkingExcelPackages
{
    class Program
    {
        static void Main(string[] args)
        {
            //epplus
            var EPPlusImport = new ImportEPPlus();
            EPPlusImport.ReadDataFromFile();
            EPPlusImport.WriteToNewFile();

            //NPOI

            //ExcelDataReader

            var summary = BenchmarkRunner.Run(typeof(Program).Assembly);
        }
    }
}
