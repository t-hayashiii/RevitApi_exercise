#region Namespaces
using System;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace UiCs
{
    [Transaction(TransactionMode.Manual)]
    public class UIInstallTest : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            TaskDialog.Show("インストールテスト", "大大成功!!");

            return Result.Succeeded;
        }
    }
}
