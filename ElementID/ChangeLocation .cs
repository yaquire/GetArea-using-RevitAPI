using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace ElementID
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class ChangeLocation : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UIDoc & Doc
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            try
            {
                Reference pickObj = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                if (pickObj != null) {
                    //Retrive Element 
                    ElementId eleID = pickObj.ElementId;
                    Element ele  = doc.GetElement(eleID);

                    using (Transaction trans = new Transaction(doc, "Change Location"))
                    {
                        
                        //Set location 
                        LocationPoint locP = ele.Location as LocationPoint;
                        if (locP != null) {
                            trans.Start();
                            XYZ loc = locP.Point;
                            XYZ newloc = new XYZ (loc.X +3, loc.Y, loc.Z);

                            locP.Point = newloc;
                            trans.Commit();
                        }

                   
                    }
                }


                return Result.Succeeded;
            }
            catch (Exception err) { 
                message = err.Message;
                return Result.Failed;
            }
        }
    }
}
