using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System;
using System.Collections.Generic;
using System.IO; // Add this line
using OfficeOpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Excel_Lib
{
    public class Excel_Read
    {
        public  List<List<string>> ReadExcel(string filepath)
        {
            var excelData = new List<List<string>>();

            // Ensure EPPlus is set to non-commercial mode (if necessary)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //ensure the file exsists
            if (!File.Exists(filepath))
            {
                Console.WriteLine("File not Found");
                return excelData;
            }
            using (ExcelPackage package = new ExcelPackage(filepath))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Get the first worksheet

                for (int row = 20; row <= 22; row++)
                {
                    List<String> col_List = new List<String>();

                    for (int col = 1; col <= 100; col++)
                    {
                        string cellValue = worksheet.Cells[row, col].Text;
                        col_List.Add(cellValue);
                    }
                    excelData.Add(col_List);
                }
            }
            return excelData;
        }
    }
}
