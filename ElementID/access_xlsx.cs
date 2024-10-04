using System;
using System.IO;
using System.Net.Http;
using ClosedXML.Excel;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ElementID
{
    internal class access_xlsx
    {
        static void main()
        {
            //Access the xlsx file, there is no imput, it is static
            using (var workbook = new XLWorkbook("C:\\Users\\yaqub\\Desktop\\ETTV-Calculation EDITED (1).xlsx"))
            {
                //access the Cover Page sheet
                var worksheet = workbook.Worksheet("CoverPage");
                //Set Range at t1-t100
                var range = worksheet.Range("T1;T100");
                //Retrives data from each cell
                foreach (var cell in range.Cells()) {
                    var value = cell.Value;
                    Console.WriteLine($"Cell {cell.Address}: {value} ");
                }
            }
        }
    }
}