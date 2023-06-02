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
using Autodesk.Revit.UI.Selection;


#endregion

namespace IntroCs
{
    [Transaction(TransactionMode.Manual)]
    public class ElementModification : IExternalCommand
    {
        // Member variables
        Application m_rvtApp;
        Document m_rvtDoc;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication rvtUIApp = commandData.Application;
            UIDocument rvtUIDoc = rvtUIApp.ActiveUIDocument;
            m_rvtApp = rvtUIApp.Application;
            m_rvtDoc = rvtUIDoc.Document;

            // 要素を選択1
            Reference refpick = rvtUIDoc.Selection.PickObject(
                ObjectType.Element, "要素を選択");
            Element elem = m_rvtDoc.GetElement(refpick);

            using (Transaction transaction = new Transaction(m_rvtDoc))   // さらにトランザクション?
            {
                // トランザクションスタート
                transaction.Start("Modify Element");    
                
                // 壁要素プロパティの修正
                ModifyElementPropertiesWall(elem);
                m_rvtDoc.Regenerate();      // グラフィックの再作図

                // 要素を選択2
                Reference refpick2 = rvtUIDoc.Selection.PickObject(
                    ObjectType.Element, "別の要素を選択");
                Element elem2 = m_rvtDoc.GetElement(refpick2);

                // 要素の移動と回転
                ModifyElementByTransformUtilsmethods(elem2);
                m_rvtDoc.Regenerate();      // グラフィックの再作図

                // トランザクションフィニッシュ
                transaction.Commit();   
            }

            return Result.Succeeded;
        }


        // ---------------------------------------------------
        // メイン関数
        // ---------------------------------------------------
        public void ModifyElementPropertiesWall(Element elem)
        {
            // 壁要素に限定
            // 定数で初期値を設定(適宜変更)

            const string wallFamilyName = "標準壁";
            const string wallTypeName = "外壁-メタル スタッド-レンガ";
            const string wallFamilyAndTypeName = 
                wallFamilyName + ": " + wallTypeName;
            const string targetLevel = "レベル 1";
            const double aboveOfset = 5000.0;
            const string targetcomment = "APIで変更";
            const double targetmove = 1000.0;


            if (!(elem is Wall))
            {
                TaskDialog.Show("Revit Intro Lab", "壁要素を選択してください");
                return;  //★重要★ これがないと先へ進んでしまう。
            }
            Wall aWall = (Wall) elem;  // 選択要素は壁

            // keep the message to the user.
            string msg = "変更された壁:" + "\n\n";

            // (1) 壁タイプを別のタイプへ変更
            // (You can enhance this to import symbol if you want.)

            Element newWallType = ElementFiltering.FindFamilyType(
                m_rvtDoc, typeof(WallType), wallFamilyName, wallTypeName, null);

            if (newWallType != null) 
            {
                aWall.WallType = (WallType) newWallType;
                msg = msg + "Wall type to: " + wallFamilyAndTypeName + "\n";
            }

            // (2) パラメータの変更
            // レベルの取得
            Level level1 = (Level)ElementFiltering.FindElement(m_rvtDoc, typeof(Level), targetLevel, null);
            if (level1 != null) 
            {
                // 「上部レベル」の変更
                aWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level1.Id);
            }
            msg += "上部レベル: " + targetLevel + "\n";

            // 「上部レベル オフセット」の変更
            double topOffset = mmToFeet(aboveOfset);  // mmをフィートへ変換メソッド
            aWall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET).Set(topOffset);

            // 「コメント」の変更
            aWall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).
                Set(targetcomment);

            msg += "上部レベル オフセット: " + aboveOfset.ToString("0") + "\n";
            msg += "コメント変更: " + targetcomment + "\n";

            // (3) 壁を(-1000, 0, 0)まで移動

            LocationCurve wallLocation = (LocationCurve)aWall.Location;
            XYZ pt1 = wallLocation.Curve.GetEndPoint(0);    // 始点
            XYZ pt2 = wallLocation.Curve.GetEndPoint(1);    // 終点

            // 移動の初期値
            double dt = mmToFeet(targetmove);
            XYZ newPt1 = new XYZ(pt1.X - dt, pt1.Y - dt, pt1.Z);
            XYZ newPt2 = new XYZ(pt2.X - dt, pt2.Y - dt, pt2.Z);

            // create a new line bound.
            Line newWallLine = Line.CreateBound(newPt1, newPt2);

            // finally change the curve.
            wallLocation.Curve = newWallLine;

            // message to the user.
            msg += "位置: 要素の始点をX方向に-1000移動" + "\n";
            TaskDialog.Show("要素の修正", msg);
        }

        public void ModifyElementByTransformUtilsmethods (Element elem)
        {
            // keep the message to the user.
            string msg = "変更された要素:" + "\n\n";

            // 1. 要素の移動
            double dt = mmToFeet(1000.0);
            // 移動したいポイント
            XYZ v = new XYZ(dt, dt, 0);

            ElementTransformUtils.MoveElement(m_rvtDoc, elem.Id, v);

            msg += "点(1000, 1000, 0)へ移動しました" + "\n";

            // 2. 要素の回転 _ z軸廻りを15度回転
            XYZ pt1 = XYZ.Zero;
            XYZ pt2 = XYZ.BasisZ;
            Line axis = Line.CreateBound(pt1, pt2);  // 始点終点から回転軸線を作成

            ElementTransformUtils.RotateElement(m_rvtDoc, elem.Id, axis, Math.PI / 12);

            msg += "Z軸を軸に15度回転しました" + "\n";

            // message to the user.
            TaskDialog.Show("要素の修正 by utils methods", msg);
        }


        // ---------------------------------------------------
        // ヘルパー関数
        // ---------------------------------------------------

        // 単位をmmからフィートに帰る単純な関数
        const double _mmToFeet = 0.0032808399;  // 変換用定数
        public static double mmToFeet(double mmValue)
        {
            return mmValue * _mmToFeet;
        }


    }
}