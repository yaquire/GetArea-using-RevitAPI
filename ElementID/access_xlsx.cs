using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

using System.IO; // Add this line
using DocumentFormat.OpenXml.Wordprocessing;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.IO.Packaging;
using System.Linq;

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
            // Open the existing Excel file
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, true))
            {
                // Get the workbook part
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                // Get the first worksheet (or select the desired one if multiple sheets exist)
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet == null)
                {
                    TaskDialog.Show("Excel","No sheet found.");
                    return;
                }
                // Get the associated worksheet part
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                // Get the SheetData where rows and cells are stored
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Add a new row and cells with data to the existing sheet
                Row row = new Row();
                sheetData.Append(row);

                row.Append(
                    CreateTextCell("A3", "Jane Doe"),
                    CreateTextCell("B3", "25"),
                    CreateTextCell("C3", "Los Angeles"));

                // Save changes to the worksheet
                worksheetPart.Worksheet.Save();
                TaskDialog.Show("Excel", "Written");
            }
        }

        private static Cell CreateTextCell(string cellReference, string text)
        {
            return new Cell()
            {
                CellReference = cellReference,
                DataType = CellValues.String,
                CellValue = new CellValue(text)
            };
        }
    }
}
    
    

