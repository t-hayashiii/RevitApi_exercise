#region Namespaces
using System;
using System.Diagnostics;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
// using Util;
using System.IO;
#endregion

namespace IntroCs
{
    [Transaction(TransactionMode.Manual)]
    public class SharedParameter : IExternalCommand
    {
        // 定数の宣言
        const string kSharedParamsGroupAPI = "API Parameters";
        const string kSharedParamsDefFireRating = "API 耐火等級";
        const string kSharedParamsPath = "C:\\temp\\SharedParams.txt";
        
        public Result Execute(
            ExternalCommandData commandData, 
            ref string message, 
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Application app = commandData.Application.Application;
            Document doc = uidoc.Document;

            // Get the current shared params definitionfile
            // 現在の共有パラメーターの定義ファイルを取得する
            DefinitionFile sharedParamsFile = GetSharedParamsFile(app);
            if (null == sharedParamsFile)
            {
                message = "Error getteing the shared params file.";
                return Result.Failed;
            }

            // Get or create the shared params group.
            // パラメータグループを取得または作成
            DefinitionGroup sharedParamsGroup = GetOrCreateSharedParamsGroup(
                sharedParamsFile, kSharedParamsGroupAPI);
            if (null == sharedParamsGroup)
            {
                message = "Error getting the shared params group.";
                return Result.Failed;
            }

            // Create the category set for binding and add the category
            // 必要とするカテゴリを取得して、共有パラメータ定義を作成
            // we are interested in, doors or walls or whatever:

            Category cat =
                doc.Settings.Categories.get_Item(BuiltInCategory.OST_Doors);

            // カテゴリーがプロジェクトパラメーターを持つことができるかどうかを確認する変数
            bool visible = cat.AllowsBoundParameters;

            // Get or create the shared params definition
            // 共有パラメーターの定義を取得または作成
            Definition fireRatingParamDef = GetOrCreateSharedParamsDefinition(
                sharedParamsGroup, ParameterType.Number, kSharedParamsDefFireRating, visible);
            if (null == fireRatingParamDef)
            {
                message = "Error in creating shared parameter.";
                return Result.Failed;   
            }

            CategorySet catSet = app.Create.NewCategorySet();
            try 
            {
                catSet.Insert(cat);     // 指定したカテゴリをセットに挿入
            }
            catch(Exception)
            {
                message = string.Format(
                    "Error adding '[0]' category to parameters binding set.", cat.Name);
                return Result.Failed;
            }

           
            using (Transaction transaction = new Transaction(doc))
            {
                transaction.Start("Bind parameter");
                // Bind tha param
                try
                {
                    Binding binding = app.Create.NewInstanceBinding(catSet);
                    doc.ParameterBindings.Insert(fireRatingParamDef, binding);
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    message = ex.Message;
                    transaction.RollBack();
                    return Result.Failed;
                }
            }
           
            return Result.Succeeded;
        }


        // ----------------------
        // ヘルパー関数
        // ----------------------

        // 共有パラメータファイルを返す関数
        // 共有パラメータが存在しない際には、新しい共有パラメータファイルを作成する機能を持つ
        public static DefinitionFile GetSharedParamsFile(Application app)
        {
            // Get current shared parames file name.
            string sharedParamsFileName;
            try
            {
                sharedParamsFileName = app.SharedParametersFilename;
            }
            catch(Exception ex) 
            {
                TaskDialog.Show(
                    "Get shared params file", "No shared params file set:" + ex.Message);
                return null;
            }

            if (0 == sharedParamsFileName.Length || !System.IO.File.Exists(sharedParamsFileName) )
            {
                StreamWriter stream;
                stream = new StreamWriter(kSharedParamsPath);
                stream.Close();
                app.SharedParametersFilename = kSharedParamsPath;
                sharedParamsFileName = app.SharedParametersFilename;
            }

            // Get the current file object and return it.
            DefinitionFile sharedPaarametersFile;
            try
            {
                sharedPaarametersFile = app.OpenSharedParameterFile();
            }

            catch (Exception ex) 
            {
                TaskDialog.Show(
                    "Get shared params file", "Cannot open shared params file:" + ex.Message);
                sharedPaarametersFile = null;
            }
            return sharedPaarametersFile    ;
        }

        // パラメータグループが既に存在するかチェックして、それを返す関数
        public static DefinitionGroup GetOrCreateSharedParamsGroup(
            DefinitionFile sharedParametersFile,
            string groupName)
        {
            DefinitionGroup g = sharedParametersFile.Groups.get_Item(groupName);
            if (g == null)
            {
                try
                {
                    g = sharedParametersFile.Groups.Create(groupName);  
                }
                catch (Exception) 
                {
                    g = null;
                }
            }
            return g;
        }

        // 与えられたパラメータ定義が既に存在するかチェックして、それを返す関数
        public static Definition GetOrCreateSharedParamsDefinition(
            DefinitionGroup defGroup,
            ParameterType defType,
            string defName,
            bool visible)
        {
            Definition definition = defGroup.Definitions.get_Item(defName);
            if (definition == null) 
            {
                try
                {
                    ExternalDefinitionCreationOptions opt
                        = new ExternalDefinitionCreationOptions(defName, defType);
                    opt.Visible = true;
                    definition = defGroup.Definitions.Create(opt);
                }
                catch (Exception) 
                {
                    definition = null;
                }
            }
            return definition;
        }

    }
}
