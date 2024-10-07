using OfficeOpenXml;
using System;
using System.Collections.Generic;

using System.IO; // Add this line
using DocumentFormat.OpenXml.Wordprocessing;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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
                for (int col = 20; col <= 22; col++)
                {
                    List<String> col_List = new List<String>();
                    for (int row = 1; row <= 10; row++)
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
    public class Write_Excel
    {
        public void writeExcel(List<List<string>> excelData, string filepath)
        {
            if (!File.Exists(filepath))
            {
                TaskDialog.Show("Error","File not Found");
                return;
            }
            // Ensure EPPlus is set to non-commercial mode (if necessary)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(filepath))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1]; // Get the first worksheet
                for (int col = 20; col <= 22; col++) {
                    for (int row = 1;row <= 10; row++)
                    {
                        //TaskDialog.Show("Col Number: Row Number", col.ToString()+":"+row.ToString());
                        try
                        {
                            worksheet.Cells[row, col].Value = excelData[0][row];
                            FileInfo fileInfo = new FileInfo(filepath);
                            package.SaveAs(fileInfo);
                        }
                        catch (Exception err)
                        {
                            break;
                        }
                        // Save the Excel package to a file
                        
                    }
                }
                
            }

        }
    }
}
