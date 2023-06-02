#region Namespaces
using System;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
// using Util;  ← ??
using IntroCs;
#endregion

namespace IntroCs
{
    [Transaction(TransactionMode.Manual)]
    public class ModelCreation : IExternalCommand
    {
        // member variables
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

            // 家の作成
            CreateHouse(m_rvtDoc);

            return Result.Succeeded;
        }


        public static void CreateHouse(Document rvtDoc)
        {
            using (Transaction transaction = new Transaction(rvtDoc))
            {
                transaction.Start("Create House");

                // 4つの壁を作成
                List<Wall> walls = CreateWalls(rvtDoc);

                // ドアの追加
                AddDoor(rvtDoc, walls[0]);

                // ★残りの壁に窓を追加
                for (int i = 1; i <=3; i++)
                {
                    AddWindow(rvtDoc, walls[i]);
                }

                transaction.Commit();
            }
        }

        public static List<Wall> CreateWalls(Document rvtDoc)
        {
            // 初期値_家のサイズ
            double width = ElementModification.mmToFeet(10000.0);
            double depth = ElementModification.mmToFeet(5000.0);

            // レベル取得
            Level level1 = (Level)ElementFiltering.FindElement(
                rvtDoc, typeof(Level), "レベル 1", null);

            if (level1 == null)
            {
                TaskDialog.Show
                    ("Revit Intro Lab", "(レベル 1)を見つけることができませんでした");
                return null;
            }

            Level level2 = (Level)ElementFiltering.FindElement(
                rvtDoc, typeof(Level), "レベル 2", null);

            if (level2 == null)
            {
                TaskDialog.Show
                    ("Revit Intro Lab", "(レベル 2)を見つけることができませんでした");
                return null;
            }

            // 壁のコーナー4点を設定
            // 5点目はループさせるために1点目と同じ点

            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> pts = new List<XYZ>(5);
            pts.Add(new XYZ(-dx, -dy, 0.0));
            pts.Add(new XYZ(dx, -dy, 0.0));
            pts.Add(new XYZ(dx, dy, 0.0));
            pts.Add(new XYZ(-dx, dy, 0.0));
            pts.Add(pts[0]);

            // flag for structural wall or not.
            bool isStructural = false;

            // save walls we create.
            List<Wall> walls = new List<Wall>(4);
            // loop through list of points and define four walls.
            for (int i = 0; i <= 3; i++)
            {
                Line baseCurve = Line.CreateBound(pts[i], pts[i + 1]);
                Wall aWall = Wall.Create(rvtDoc, baseCurve, level1.Id, isStructural);
                // set the Top Constraint to Level 2
                aWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id);
                // save the wall
                walls.Add(aWall);
            }

            // ★重要★ we need these lines to have shrinkwrap working.
            rvtDoc.Regenerate();
            rvtDoc.AutoJoinElements();

            return walls;
        }

        public static void AddDoor(Document rvtDoc, Wall hostWall)
        {
            // 初期値の定数
            const string doorFamilyName = "片開き";
            const string doorTypeName = "w800h2000";
            const string doorFamilyAndTypeName = 
                doorFamilyName + ": " + doorTypeName;

            // ドアタイプの取得
            FamilySymbol doorType = 
                (FamilySymbol) ElementFiltering.FindFamilyType(
                    rvtDoc, typeof(FamilySymbol), doorFamilyName, doorTypeName,
                    BuiltInCategory.OST_Doors);

            if (doorType == null) 
            {
                TaskDialog.Show("Revit Intro Lab", 
                    doorFamilyAndTypeName + "を見つけられませんでした");
            }

            // ホスト壁の始点と終点を取得
            LocationCurve locCurve = (LocationCurve)hostWall.Location;
            XYZ pt1 = locCurve.Curve.GetEndPoint(0);
            XYZ pt2 = locCurve.Curve.GetEndPoint(1);
            //中点の計算
            XYZ pt = (pt1 + pt2) / 2.0;

            // ホスト壁からレベルを取得
            ElementId idLevel1 = hostWall.get_Parameter(
                BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId();
            Level level1 = (Level) rvtDoc.GetElement(idLevel1);

            // ドアを作成
            FamilyInstance aDoor =
                rvtDoc.Create.NewFamilyInstance(
                    pt, doorType, hostWall, level1, StructuralType.NonStructural);
        }
    
        public static void AddWindow(Document rvtDoc, Wall hostWall)
        {
            // 初期値
            const string windowFamilyName = "横すべり窓";
            const string windowTypeName = "w1000h0400";
            const string windowFamilyAndTypeName =
                windowFamilyName + ": " + windowTypeName;
            double sillHeight = ElementModification.mmToFeet(915);

            // 窓タイプの取得
            FamilySymbol windowType = 
                (FamilySymbol) ElementFiltering.FindFamilyType(
                    rvtDoc, typeof(FamilySymbol), windowFamilyName, windowTypeName,
                    BuiltInCategory.OST_Windows);

            if (windowType == null) 
            {
                TaskDialog.Show("Revit Intro Lab", windowFamilyAndTypeName + "を見つけることができません");
            }

            // ホスト壁の始点と終点を取得
            LocationCurve locCurve = (LocationCurve)hostWall.Location;
            XYZ pt1 = locCurve.Curve.GetEndPoint(0);
            XYZ pt2 = locCurve.Curve.GetEndPoint(1);
            //中点の計算
            XYZ pt = (pt1 + pt2) / 2.0;

            // ホスト壁からレベルを取得
            ElementId idLevel1 = hostWall.get_Parameter(
                BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId();
            Level level1 = (Level)rvtDoc.GetElement(idLevel1);

            // 窓を作成
            FamilyInstance aWindow = rvtDoc.Create.NewFamilyInstance(
                pt, windowType, hostWall, level1, StructuralType.NonStructural);

            // 腰高パラメータの設定
            aWindow.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).
                Set(sillHeight);
        }
    }
}
