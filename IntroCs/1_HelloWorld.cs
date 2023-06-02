#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

#endregion

namespace IntroCs
{
    [Transaction(TransactionMode.ReadOnly)]
   
    /// Hello World #1 - A minimum revit external command.
    public class HelloWorldSimple : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData revit,
            ref string message, 
            ElementSet elements)
        {
            TaskDialog.Show("My Dialog Title", "Hello World Simple!");
            return Result.Succeeded;
        }
    }

    // Hello World #3 - minimum external application
    public class HelloWorldApp : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            TaskDialog.Show("My Dialog Title", "Hello World from App!");
            return Result.Succeeded;
        }
    }

    // Command Arguments and Revit Object Model
    [Transaction(TransactionMode.ReadOnly)]
    public class CommandData : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData, 
            ref string message, 
            ElementSet elements)
        {
            // The first argument, commandData, provides access to the top most object model.
            // You will get the necessary information from commandData.
            // To see what's in there, print out a few data accessed from commandData
            // Exercise: Place a break point at commandData and drill down the data.

            UIApplication rvtUiApp = commandData.Application;
            Application rvtApp = rvtUiApp.Application;
            UIDocument rvtUiDoc = rvtUiApp.ActiveUIDocument;
            Document rvtDoc = rvtUiDoc.Document;

            // Print out a few information thai you can get from commandData
            string versionName = rvtApp.VersionName;
            string documentTitle = rvtDoc.Title;

            TaskDialog.Show(
                "Revit Intro Lab",
                "Version Name = " + versionName
                + "\nDocument Title = " + documentTitle);

            // Print out a list of wall types available in the current rvt project:

            FilteredElementCollector collector = new FilteredElementCollector(rvtDoc);
            collector.OfClass(typeof(WallType));

            string s = "";
            foreach(WallType wallType in collector)
            {
                s += wallType.Name + "\r\n";
            }

            // Show the result:
            TaskDialog.Show(
                "Revit Intro Lab",
                "Wall Types (in main instruction):\n\n" + s);

            // 2nd and 3rd arguments are when the command fails.
            // 2nd - set a message to the user.
            // 3rd - set elements to highlight.

            return Result.Succeeded;
        }
    }
}
