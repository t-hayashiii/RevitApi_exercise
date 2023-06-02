#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

#endregion

namespace IntroCs
{
    [Transaction(TransactionMode.Manual)]
    public class ElementFiltering : IExternalCommand
    {
        // Member variables
        Application m_rvtApp;
        Document m_rvtDoc;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            // Get the access to the top most objects.
            UIApplication rvtUIApp = commandData.Application;
            UIDocument rvtUIDoc = rvtUIApp.ActiveUIDocument;
            m_rvtApp = rvtUIApp.Application;
            m_rvtDoc = rvtUIDoc.Document;

            // (1)ファミリタイプをリスト化
            ListFamilyTypes();

            // (2)インスタンスをリストで取得
            ListInstances();

            // (3)特定のファミリタイプを見つける 
            ElementType wallType3 =
                (ElementType)FindFamilyType(m_rvtDoc, typeof(WallType), 
                "標準壁", "sda", null);
            ElementType doorType3 =
                (ElementType)FindFamilyType(m_rvtDoc, typeof(FamilySymbol),
                "片開き", "w900h2000", BuiltInCategory.OST_Doors);

            string msg = "特定のファミリタイプ: \r\n\r\n" 
                + wallType3.Name + "\n"
                + doorType3.Name + "\n";          
            TaskDialog.Show("特定のファミリタイプ", msg);

            // (4)特定インスタンスを見つける
            Level level1 = (Level)FindElement(m_rvtDoc, typeof(Level), "レベル 1", null);
            TaskDialog.Show("特定のインスタンス", level1.Name);


            return Result.Succeeded;
        }


        //====================================================================
        // メイン関数
        //====================================================================

        // ファミリタイプをリストで取得
        public void ListFamilyTypes()
        {
            // (1) Get a list of family types available in the current project.
            //
            // for system family types, there is a designated.
            // properties that allows us to directly access to the types.
            // e.g., _doc.WallTypes

            // WallTypeSet wallTypes = _doc.WallTypes;  // 2013
            // 壁タイプのフィルタリング

            FilteredElementCollector wallTypes
                = new FilteredElementCollector(m_rvtDoc)
                .OfClass(typeof(WallType));
            int n = wallTypes.Count();
            string s = string.Empty;

            foreach (WallType wType in wallTypes)
            {
                s += "\r\n" + wType.Kind.ToString() + " : " + wType.Name;
            }
            TaskDialog.Show(n.ToString() + " Wall Types:", s);

            // (1.1) Same idea applies to other system family, such as Floors, Roofs.
            // FloorTypeSet floorTypes = _doc.FloorTypes;
            // 床タイプのフィルタリング

            FilteredElementCollector floorTypes
                = new FilteredElementCollector(m_rvtDoc)
                .OfClass(typeof(FloorType));
            int m = floorTypes.Count();
            s = string.Empty;

            foreach (FloorType fType in floorTypes)
            {
                // Family name is not in the property for floor,
                // so use BuiltInParameter here.

                Parameter param = fType.get_Parameter(
                    BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM);

                if (param != null)
                {
                    s += param.AsString();
                }
                s += " : " + fType.Name + "\r\n";
            }
            TaskDialog.Show(m.ToString() + " floor Types:", s);

            // (1.2a) Another approach is to use a filter, here is an example with wall type.
            // 壁タイプフィルタリングの別のアプローチ (シンプル案?)

            var wallTypeCollector1 = new FilteredElementCollector(m_rvtDoc);
            wallTypeCollector1.WherePasses(new ElementClassFilter(typeof(WallType)));
            IList<Element> wallTypes1 = wallTypeCollector1.ToElements();

            // 下記で定義したメソッドを使用してshow
            ShowElementList(wallTypes1, "Wall Types (by Filter): ");


            // (2) Listing for component family types.
            // 
            // For component family, it is slightly different.
            // There is no designate property in the document class.
            // You always need to use a filtering. e.g., for doors and windows.
            // Remember for component family, you will need to check element type and category.
            // コンポーネントファミリのフィルタリング(システムファミリ以外)

            var doorTypeCollector = new FilteredElementCollector(m_rvtDoc);
            doorTypeCollector.OfClass(typeof(FamilySymbol));
            doorTypeCollector.OfCategory(BuiltInCategory.OST_Doors);
            IList<Element> doorTypes = doorTypeCollector.ToElements();

            ShowElementList(doorTypes, "Door Types (by Filter): ");
        }

        // インスタンスをリストで取得
        public void ListInstances()
        {
            // 全ての壁インスタンスをリストで取得(システムファミリ)
            var wallCollector = new FilteredElementCollector(m_rvtDoc).OfClass(typeof(Wall));
            IList<Element> wallList = wallCollector.ToElements();

            ShowElementList(wallList, "Wall Instances: ");

            // 全てのドアインスタンスをリストで取得(システムファミリ以外)
            var doorCollector = new FilteredElementCollector(m_rvtDoc).OfClass(typeof(FamilyInstance));
            doorCollector.OfCategory(BuiltInCategory.OST_Doors);
            IList<Element> doorList = doorCollector.ToElements();

            ShowElementList(doorList, "Door Instance: ");
        }

        // 特定の名前(文字列)と一致する壁タイプを見つける ~汎用的な定義~
        public static Element FindFamilyType(
            Document rvtDoc, 
            Type targetType,
            string targetFamilyName, 
            string targetTypeName, 
            Nullable<BuiltInCategory> targetCategory)
        {
            // First, narrow down to the elements of the given type and category 
            var collector = new FilteredElementCollector(rvtDoc).OfClass(targetType);
            if (targetCategory.HasValue)
            {
                collector.OfCategory(targetCategory.Value);
            }

            // Parse the collection for the given names 
            // Using LINQ query here. 
            var targetElems =
                from element in collector
                where element.Name.Equals(targetTypeName) &&
                element.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString().Equals(targetFamilyName)
                select element;

            // Put the result as a list of element fo accessibility. 
            IList<Element> elems = targetElems.ToList();

            // Return the result. 
            if (elems.Count > 0)
            {
                return elems[0];
            }
            return null;
        }

        // 5.1 特定のファミリタイプに該当するインスタンスを見つける
        // find a list of element with given class, family type and category(optional).
        public IList<Element> FindInstanceOfType(
            Type targetType,
            ElementId idType,
            Nullable<BuiltInCategory> targetCategory
            )
        {
            // narrow down to the elements of the given type and category

            var collector =
                new FilteredElementCollector(m_rvtDoc).OfClass(targetType); 
            if (targetCategory.HasValue) 
            {
                collector.OfCategory(targetCategory.Value);
            }

            // parse the collection for the given family type id. using LINQ query here.
            // LINQクエリを使用して、与えられたファミリータイプIDのコレクションをパースします。

            var elems =
                from element in collector
                where element.get_Parameter(BuiltInParameter.SYMBOL_ID_PARAM).AsElementId().Equals(idType)
                select element;

            return elems.ToList();
        }

        // 5.2.1 与えられたクラスと名前で要素を見つける_リストで取得
        public static IList<Element> FindElements (
            Document rvtDoc,
            Type targetType,
            string targetName,
            Nullable<BuiltInCategory> targetCategory)
        {
            var collector =
                new FilteredElementCollector(rvtDoc).OfClass(targetType);
            if (targetCategory.HasValue)
            {
                collector.OfCategory(targetCategory.Value);
            }

            var elems =
                from element in collector
                where element.Name.Equals(targetName)
                select element;

            return elems.ToList();
        }

        // 5.2.2 与えられたクラスと名前で要素を見つける_ひとつだけ取得
        public static Element FindElement(
            Document rvtDoc,
            Type targetType,
            string targetName,
            Nullable<BuiltInCategory> targetCategory)
        {
            IList<Element> elems =
                FindElements(rvtDoc, targetType, targetName, targetCategory);

            // return the first one from the result.
            if (elems.Count > 0)
            {
                return elems[0];
            }
            return null;
        }


        //====================================================================
        // Helper Functions ヘルパー関数
        //====================================================================
        // ShowElementListメソッドの定義
        public void ShowElementList(IList<Element> elems, string header)
        {
            string s = " - Class - Category - Name (or Family: Type Name) - Id - \r\n";
            foreach (Element e in elems)
            {
                s += ElementToString(e);
            }
            TaskDialog.Show(header + "(" + elems.Count.ToString() + "):", s);
        }

        // ElementToStringメソッドの定義
        public string ElementToString(Element e)
        {
            if (e == null)
            {
                return "none";
            }

            string name = "";
            if (e is ElementType)
            {
                Parameter param = e.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM);
                if (param != null)
                {
                    name = param.AsString();
                }
            }
            else
            {
                name = e.Name;
            }

            return e.GetType().Name + ";"
                + e.Category.Name + ";"
                + name + ";"
                + e.Id.IntegerValue.ToString() + "\r\n";
        }

        /* 
        ShowFamilyTypeAndIdメソッドの定義
        public string ShowFamilyTypeAndId
            (string header, ElementType familyType)
        {
            // Show the result. 
            string msg = header + "\r\n" ;

            if (familyType != null)
            {
                msg += familyType.Name + "\r\n";
            }
             
            return msg;
        }
         */


    }
}
