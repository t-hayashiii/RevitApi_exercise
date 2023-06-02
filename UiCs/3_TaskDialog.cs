#region Namespaces
using System;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
// using Util;
#endregion

namespace UiCs
{
    [Transaction(TransactionMode.Manual)]
    public class UITaskDialog : IExternalCommand
    {
        UIApplication _uiApp;
        UIDocument _uiDoc;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            _uiApp = commandData.Application;
            _uiDoc = _uiApp.ActiveUIDocument;

            // (1)
            // ShowTaskDialogStatic();

            // (2)
            ShowTaskDialogInstance(true);

            return Result.Succeeded;
        }


        /// ShowTaskDialogStatic()
        /// 
        public void ShowTaskDialogStatic()
        {
            // (1) シンプルなダイアログ
            TaskDialog.Show("Task Dialog Static 1", "Main message");

            // (2) this version accepts command buttons in addition to above. 
            // Here we add [Yes] [No] [Cancel} 
            TaskDialogResult res2 = default(TaskDialogResult);
            res2 = TaskDialog.Show("Task Dialog Static 2", "Main message",
                (TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No | TaskDialogCommonButtons.Cancel));

            // ユーザーが押したボタンの表示
            TaskDialog.Show("Show task dialog", "押したボタン: " + res2.ToString());

            // (3) this version accepts default button in addition to above. 
            // Here we set [No] as a default (just for testing purposes). 
            // デフォルト選択の追加
            TaskDialogResult res3 = default(TaskDialogResult);
            TaskDialogResult defaultButton = TaskDialogResult.No;
            res3 = TaskDialog.Show("Task Dialog Static 3", "Main message",
                (TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No | TaskDialogCommonButtons.Cancel),
                defaultButton);

            // ユーザーが押したボタンの表示
            TaskDialog.Show("Show task dialog", "押したボタン: " + res3.ToString());

        }

        /// ShowTaskDialogInstance
        /// 
        public void ShowTaskDialogInstance(bool stepByStep)
        {
            // (0) create an instance of task dialog to set more options. 
            // (0) タスクダイアログのインスタンスを作成し、より多くのオプションを設定します。

            TaskDialog myDialog = new TaskDialog("Revit UI Labs - Task Dialog Options");
            if (stepByStep) myDialog.Show();  // Just declare stepBystep stepBystepを宣言するだけ


            // (1) set the main area. These appear at the upper portion of the dialog.
            // (1) メインエリアの設定。これらはダイアログの上部に表示される。

            myDialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
            // or TaskDialoIcon.TaskDialogIconNone.
            if (stepByStep) myDialog.Show();

            myDialog.MainInstruction =
                "Main instruction: This is Revit UI Lab 3 Task Dialog";
            if (stepByStep) myDialog.Show();

            myDialog.MainContent =
                "Main content: You can add detailed description here.";
            if (stepByStep) myDialog.Show();


            // (2) set the bottom area
            myDialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No | TaskDialogCommonButtons.Cancel;
            myDialog.DefaultButton = TaskDialogResult.Yes;
            if (stepByStep) myDialog.Show();

            myDialog.ExpandedContent = "Expanded content: the visibility of this portion is controled by Show/Hide button.";
            if (stepByStep) myDialog.Show();

            myDialog.VerificationText = "Verification: Do not show this message again comes here";
            if (stepByStep) myDialog.Show();

            myDialog.FooterText = "Footer: <a href=\"http://www.autodesk.com/developrevit\">Revit Developer Center</a>";
            if (stepByStep) myDialog.Show();


            // (4) add command links. you can add up to four links 
            myDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Commanad Link 1", "description 1");
            if (stepByStep) myDialog.Show();

            myDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Commanad Link 2", "description 2");
            if (stepByStep) myDialog.Show();

            myDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Commanad Link 3",
                "you can add up to four commnad links");
            if (stepByStep) myDialog.Show();

            myDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink4, "Commanad Link 4", 
                "Can also have URLs e.g. Revit Online Help");
            // if (stepByStep) myDialog.Show();

            // Show it.
            TaskDialogResult res = myDialog.Show();
            if (TaskDialogResult.CommandLink4 == res)
            {
                System.Diagnostics.Process process =
                    new System.Diagnostics.Process();

                process.StartInfo.FileName =
                    "https://help.autodesk.com/view/RVT/2021/JPN/";
                process.Start();
            }

            TaskDialog.Show("Show task dialog", "The last action was: " + res.ToString());

        }
    }
}
