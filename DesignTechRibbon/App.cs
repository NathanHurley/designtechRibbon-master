using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using Autodesk.Windows;

using UIFramework;
using RibbonItem = Autodesk.Revit.UI.RibbonItem;
using VCButtonsWithVideoToolTip;
using System.IO;
using System.Media;
using static System.Environment;

namespace DesignTechRibbonPaid
{

    class App : IExternalApplication
    {

        public static string dirWin7 = Environment.GetEnvironmentVariable("UserProfile") + "\\AppData\\Roaming\\";
        public Result OnStartup(UIControlledApplication application)
        {

            // Call our method that will load up our toolbar
            AddRibbonPanel(application);


            return Result.Succeeded;
        }



        // Define a method that will create our tab and button
        static void AddRibbonPanel(UIControlledApplication application)
        {
            // Create a custom ribbon tab
            String tabName = "designtech.io (Paid)";
            application.CreateRibbonTab(tabName);

     
            //Autodesk.Revit.UI.RibbonPanel aboutPanel = application.CreateRibbonPanel(tabName, "About");
            Autodesk.Revit.UI.RibbonPanel toolsPanel = application.CreateRibbonPanel(tabName, "Tools");
            //Autodesk.Revit.UI.RibbonPanel viewsPanel = application.CreateRibbonPanel(tabName, "Views");
            //Autodesk.Revit.UI.RibbonPanel importPanel = application.CreateRibbonPanel(tabName, "Import");
            //Autodesk.Revit.UI.RibbonPanel exportPanel = application.CreateRibbonPanel(tabName, "Export");


            //info
            //CreateInfoButton(application, tabName, aboutPanel);

            //Tools
            //CreateAdvancedSharedParamaterMapButton(application, tabName, toolsPanel);
            //CreateAutoDimCurtainWallsButton(application, tabName, toolsPanel);
            //CreateElementsInRoomsButton(application, tabName, toolsPanel);
            CreateExportCategoryParametersButton(application, tabName, toolsPanel);

            //Views
            //CreateSelectionFilterButton(application, tabName,viewsPanel);
            //CreatePartPlansViewsButton(application, tabName, viewsPanel);
            //CreateInternalElevationsButton(application, tabName, viewsPanel);
            //CreateViewExtenderButton(application, tabName, viewsPanel);

            //Import
            //CreateRhinoToRevitButton(application, tabName, importPanel);

            //Export
            //CreateIssueModelButton(application, tabName, exportPanel);


        }

        //Manage Tab Buttons

        static void CreateInfoButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
  

            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdInfo",
               "Info", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Licence Information";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/about.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;


        }


        static void CreateAdvancedSharedParamaterMapButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdInfo",
               "Paramater\nMap", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Advanced Shared Paramater Map";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/parameter_map.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        }


        static void CreateAutoDimCurtainWallsButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdAutoDimCurtainWalls",
               "Auto-Dim\nCurtain Walls", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Auto Dimension Curtain Walls";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/auto_dimension_curtain_walls.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        }

        static void CreateElementsInRoomsButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdElementsInRooms",
               "Elements\nIn Spaces ", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Elements In Rooms";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/elements_in_spaces.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        }

        static void CreateExportCategoryParametersButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp) //#########################
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdExportCategoryParameter",
               "Export\nParameters", thisAssemblyPath, "EssentialTools.ExportCategoryParameters"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Export Category Parameter";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/export_category_parameters.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        }

        static void CreateSelectionFilterButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdSelectionFilter",
               "Selection\nFilter", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Selection Filter";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/selection_filter.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        }

        static void CreatePartPlansViewsButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdPartPlans",
               "Part-Plans", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Create Part Plans";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/create_part_plans.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        } ///which icon?

        static void CreateInternalElevationsButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdInternalElevations",
               "Internal\nElevations", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Internal Elevations";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/create_internal_elevations.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        }

        static void CreateViewExtenderButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdViewExtender",
               "View\nExtender", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "View Extender";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/viewport_extender.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        }

        static void CreateRhinoToRevitButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdRhinoToRevit",
               "Rhino\nTo Revit", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Rhino To Revit";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/rhino_to_revit.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        }

        static void CreateIssueModelButton(UIControlledApplication a, string tabname, Autodesk.Revit.UI.RibbonPanel rp)
        {
            // Create a push button to trigger a command add it to the ribbon panel.
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData("cmdIssueModel",
               "Issue\nModel", thisAssemblyPath, "EssentialTools.Info"); //Make sure this class exists

            PushButton pushButton = rp.AddItem(buttonData) as PushButton;

            pushButton.ToolTip = "Issue Model";

            Uri uriImage = new Uri((@"pack://application:,,,/DesignTechRibbonPaid;component/Resources/EssentialTools/issue_model.png"));
            // Apply image to bitmap
            BitmapImage largeImage = new BitmapImage(uriImage);
            // Apply image to button 
            pushButton.LargeImage = largeImage;
        }





        /// <summary>
        /// A method that allows you to create a Push Button with greater ease
        /// </summary>
        /// <param name="ribbonPanel"></param>
        /// <param name="name"></param>
        /// <param name="path"></param>
        /// <param name="command"></param>
        /// <param name="tooltip"></param>
        /// <param name="icon"></param>
        private static void CreatePushButton(Autodesk.Revit.UI.RibbonPanel ribbonPanel, string name, string path, string command, string tooltip, string icon)
        {
            PushButtonData pbData = new PushButtonData(
                name,
                name,
                path,
                command);

            PushButton pb = ribbonPanel.AddItem(pbData) as PushButton;
            pb.ToolTip = tooltip;
            BitmapImage pb2Image = new BitmapImage(new Uri(String.Format("pack://application:,,,/DesignTechRibbonPaid;component/Resources/{0}", icon)));
            pb.LargeImage = pb2Image;
        }

        /////Ribbon

        public Autodesk.Windows.RibbonItem GetRibbonItem(Autodesk.Revit.UI.RibbonItem item)
        {
            Type itemType = item.GetType();

            var mi = itemType.GetMethod("getRibbonItem",
              BindingFlags.NonPublic | BindingFlags.Instance);

            var windowRibbonItem = mi.Invoke(item, null);

            return windowRibbonItem
              as Autodesk.Windows.RibbonItem;
        }

        static void SetRibbonItemToolTip(RibbonItem item, RibbonToolTip toolTip)
        {
            IUIRevitItemConverter itemConverter =
                new InternalMethodUIRevitItemConverter();

            var ribbonItem = itemConverter.GetRibbonItem(item);
            if (ribbonItem == null)
                return;
            ribbonItem.ToolTip = toolTip;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            // Do nothing
            return Result.Succeeded;
        }



    }




}


