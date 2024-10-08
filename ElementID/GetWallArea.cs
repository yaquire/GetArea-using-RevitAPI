//For Revit Api 
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML;

//For Excel
using Excel_Lib;


using System;
using System.Collections.Generic;
using System.Globalization;
using System.Diagnostics; // Required for Process
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ElementID
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class GetWallArea : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //get UIdoc & doc
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uiDoc.Document;
            //TaskDialog.Show("To use", "Select Item it Will be added to a file");
            string filePath = @"C:\Users\yaqub\Desktop\New Microsoft Excel Worksheet.xlsx";
            //Read from the excel
            //Excel_Read read_Excel = new Excel_Read();
            List<List<string>> excel_data =new List<List<string>>();
            TaskDialog.Show("Data from the Excel", excel_data.Count.ToString());
            try
            {
                Reference pickObj = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                if (pickObj != null)
                {
                    //retrives the element
                    ElementId eleID = pickObj.ElementId;
                    Element ele = doc.GetElement(eleID);
                    Parameter areaParam = ele.LookupParameter("Area");

                    List<string> rowData = new List<string> ();
                    //This line gets the area of the selected item to 3dp
                    double area = areaParam.AsDouble();
                    area = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareMeters); //Converts to square meters no whether if sf or m2
                    rowData.Add("NIL");
                    rowData.Add(ele.Name.ToString());
                    rowData.Add(area.ToString());

                    excel_data.Add(rowData);
                    TaskDialog.Show("Area: ", area.ToString("F2", CultureInfo.InvariantCulture) + "_m2");
                }
                
                //Write to the excel
                Write_Excel writeExcel = new Write_Excel();
                writeExcel.writeExcel(excel_data, filePath);
                return Result.Succeeded;
            }
            catch (Exception err)
            {
                message = err.Message;
                return Result.Failed;
            }
        }  
    }
}
