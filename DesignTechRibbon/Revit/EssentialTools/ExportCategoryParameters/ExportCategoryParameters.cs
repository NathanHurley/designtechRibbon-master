using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using System.Reflection;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;
using System.Windows.Forms;
using DesignTechRibbonPaid.Revit.EssentialTools.ExportCategoryParametersForm;

namespace EssentialTools
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class ExportCategoryParameters : IExternalCommand
    {
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit, ref string message, ElementSet elements)
        {
            UIApplication uiapp = revit.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            ExportCategoryParametersForm form = new ExportCategoryParametersForm(doc);  //Calls the internal class which then calls the form

            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.ShowDialog();
            form.Dispose();
            return Result.Succeeded;
        }
    }
}