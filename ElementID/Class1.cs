using System;


using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ElementID
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            //Gets the UIDocument
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;

            try
            {
                //Pick obj
                Reference pickObj = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);

                if (pickObj != null)
                {
                    TaskDialog.Show("Element ID", pickObj.ElementId.ToString());
                }
                //TaskDialog.Show("Title", "Hello World
                return Result.Succeeded;
            }
            catch (Exception ex) { 
                message = ex.Message;
                return Result.Failed;
            }
            
        }
    }
}
