#region Namespaces

using System;
using System.Collections.Generic;
using System.Linq; // in System.Core 
using Autodesk.Revit;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using FamilyLabsCS;

#endregion // Namespaces

namespace FamilyLabsCs
{
    [Transaction(TransactionMode.Manual)]
    class RvtCmd_FamilyCreateColumnRectangle : IExternalCommand
    {
        Application _rvtApp;
        Document _rvtDoc;

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            _rvtApp = commandData.Application.Application;
            _rvtDoc = commandData.Application.ActiveUIDocument.Document;


            // (0) ファミリテンプレートのチェック実装
            if (!isRightTemplate(BuiltInCategory.OST_Columns))
            {
                Util.ErrorMsg("Please open 柱(メートル単位).rft");
                return Result.Failed;
            }
            else
            {
                TaskDialog.Show("ファミリテンプレートのチェック", "OK");
            }


            using (Transaction transaction = new Transaction(_rvtDoc))
            {
                try
                {
                    if (transaction.Start("CreateFamily") == TransactionStatus.Started)
                    {
                        // (1) 押出しの作成 
                        Extrusion pSolid = createSolid();

                        // We need to regenerate so that we can build on this new geometry
                        _rvtDoc.Regenerate();

                        // try this:
                        // if you comment addAlignment and addTypes calls below and execute only up to here,
                        // you will see the column's top will not follow the upper level.

                        // (2) 位置合わせの追加
                        addAlignments(pSolid);

                        // try this: at each stage of adding a function here, you should be able to see the result in UI.

                        // (3) タイプの追加
                        //addTypes();

                        transaction.Commit();
                    }
                    else
                    {
                        TaskDialog.Show("ERROR", "Start transaction failed!");
                        return Result.Failed;
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("ERROR", ex.ToString());
                    if (transaction.GetStatus() == TransactionStatus.Started)
                        transaction.RollBack();
                    return Result.Failed;
                }
            }

            // finally, return
            return Result.Succeeded;
        }




        bool isRightTemplate(BuiltInCategory targetCategory)
        {
            // This command works in the context of family editor only.
            //このコマンドは、ファミリーエディターのコンテキストでのみ機能
            // ファミリドキュメントかどうかをチェック
            if (!_rvtDoc.IsFamilyDocument)
            {
                Util.ErrorMsg("This command works only in the family editor.");
                return false;
            }

            // Check the template for an appropriate category here if needed.
            // 必要であれば、ここで適切なカテゴリのテンプレートを確認してください。
            // テンプレートが柱ファミリ定義用かどうかをチェック
            Category cat = _rvtDoc.Settings.Categories.get_Item(targetCategory);
            if (_rvtDoc.OwnerFamily == null)
            {
                Util.ErrorMsg("This command only works in the family contextr.");
                return false;
            }

            if (!cat.Id.Equals(_rvtDoc.OwnerFamily.FamilyCategory.Id))
            {
                Util.ErrorMsg
                    ("Category of this family document does not match the context required by this command.");
                return false;
            }

            // if we come here, we should have a right one.
            return true;
        }

        Extrusion createSolid()
        {
            CurveArrArray pProfile = createProfileRectangle();

            // スケッチ平面=作業面の作成
            ReferencePlane pRefPlane = findElement (typeof(ReferencePlane), "参照面") as ReferencePlane;
            SketchPlane pSketchPlane = SketchPlane.Create(_rvtDoc, pRefPlane.GetPlane());

            // 押出し高さ
            double dHeight = mmToFeet(4000.0);

            // 押出し
            bool bIsSolid = true;  // 常にtrue。 このメソッドの押出し用。
            return _rvtDoc.FamilyCreate.NewExtrusion(bIsSolid, pProfile, pSketchPlane, dHeight);
        }

        CurveArrArray createProfileRectangle()
        {
            double w = mmToFeet(600.0);
            double d = mmToFeet(600.0);
            const int nVerts = 4;   // 辺の数

            // 四角形の作成
            XYZ[] pts = new XYZ[]
            {
                new XYZ(-w/2.0, -d/2.0, 0.0),
                new XYZ(w/2.0, -d/2.0, 0.0),
                new XYZ(w/2.0, d/2.0, 0.0),
                new XYZ(-w/2.0, d/2.0, 0.0),
                new XYZ(-w/2.0, -d/2.0, 0.0)
            };

            CurveArray pLoop = _rvtApp.Create.NewCurveArray();
            for(int i = 0; i < nVerts; ++i)
            {
                Line line = Line.CreateBound(pts[i], pts[i + 1]);
                pLoop.Append(line);
            }

            // カーブからプロファイルを作成
            CurveArrArray pProfile = _rvtApp.Create.NewCurveArrArray();
            pProfile.Append(pLoop);

            return pProfile;
        }

        void addAlignments(Extrusion pBox)
        {
            // (1) 柱の上面を上参照レベルに拘束
            View pView = findElement(typeof(View), "正面") as View;

            Level upperLevel = findElement(typeof(Level), "上参照レベル") as Level;
            Reference ref1 = upperLevel.GetPlaneReference(); // レベルから参照面を取得

            PlanarFace upperFace = findFace(pBox, new XYZ(0, 0, 1));
            Reference ref2 = upperFace.Reference;

            // 位置合わせの作成
            _rvtDoc.FamilyCreate.NewAlignment(pView, ref1, ref2);

            // (2) 下端レベルも同じように拘束
            Level lowerLevel = findElement(typeof(Level),"下参照レベル") as Level;
            Reference ref3 = lowerLevel.GetPlaneReference();

            PlanarFace lowerFace = findFace(pBox, new XYZ(0, 0, -1));
            Reference ref4 = lowerFace.Reference;

            // 位置合わせの作成
            _rvtDoc.FamilyCreate.NewAlignment(pView, ref3, ref4);

            // (3) 同様に右左前後の拘束も行う
            View pViewPlan = findElement(typeof(ViewPlan), "下参照レベル") as View;

            // find reference planes
            ReferencePlane refRight = findElement(typeof(ReferencePlane),"右") as ReferencePlane;
            ReferencePlane refLeft  = findElement(typeof(ReferencePlane), "左") as ReferencePlane;
            ReferencePlane refFront = findElement(typeof(ReferencePlane), "正面") as ReferencePlane;
            ReferencePlane refBack  = findElement(typeof(ReferencePlane), "背面") as ReferencePlane;

            // find the face of the box
            PlanarFace faceRight = findFace(pBox, new XYZ(1.0, 0.0, 0.0));
            PlanarFace faceLeft = findFace(pBox, new XYZ(-1.0, 0.0, 0.0));
            PlanarFace faceFront = findFace(pBox, new XYZ(0.0, -1.0, 0.0));
            PlanarFace faceBack = findFace(pBox, new XYZ(0.0, 1.0, 0.0));

            // 位置合わせの作成
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refRight.GetReference(), faceRight.Reference);
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refLeft.GetReference(), faceLeft.Reference);
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refFront.GetReference(), faceFront.Reference);
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refBack.GetReference(), faceBack.Reference);

        }

        void addTypes()
        {
            addType("600x900", 600, 900);
            addType("1000x300", 1000, 300);
            addType("600x600", 600, 600);
        }

        void addType(string name, double w, double d) 
        {
            // get the family manager from the current doc
            FamilyManager pFamilyMgr = _rvtDoc.FamilyManager;

            // add new types with the given name
            FamilyType type1 = pFamilyMgr.NewType(name);

            // 幅パラメータの設定
            FamilyParameter paramW = pFamilyMgr.get_Parameter("幅");
            double valW = mmToFeet(w);
            if (paramW != null) 
            {
                pFamilyMgr.Set(paramW, valW);
            }

            // 奥行パラメータの設定
            FamilyParameter paramD = pFamilyMgr.get_Parameter("奥行き");
            double valD = mmToFeet(d);
            if (paramD != null)
            {
                pFamilyMgr.Set(paramD, valD);
            }
        }


        double mmToFeet(double mmVal)
        {
            return mmVal / 304.8;
        }

        Element findElement(Type targetType, string targetName)
        {
            // get the elements of the given type
            //
            FilteredElementCollector collector = new FilteredElementCollector(_rvtDoc);
            collector.WherePasses(new ElementClassFilter(targetType));

            // parse the collection for the given name
            // using LINQ query here. 
            // 
            var targetElems = from element in collector where element.Name.Equals(targetName) select element;
            List<Element> elems = targetElems.ToList<Element>();

            if (elems.Count > 0)
            {  // we should have only one with the given name. 
                return elems[0];
            }

            // cannot find it.
            return null;
        }

        PlanarFace findFace(Extrusion pBox, XYZ normal)
        {
            // get the geometry object of the given element
            //
            Options op = new Options();
            op.ComputeReferences = true;
            GeometryElement geomElem = pBox.get_Geometry(op);

            // loop through the array and find a face with the given normal
            //
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid) // solid is what we are interested in.
                {
                    Solid pSolid = geomObj as Solid;
                    FaceArray faces = pSolid.Faces;
                    foreach (Face pFace in faces)
                    {
                        PlanarFace pPlanarFace = pFace as PlanarFace;
                        if ((pPlanarFace != null) && pPlanarFace.FaceNormal.IsAlmostEqualTo(normal)) // we found the face
                        {
                            return pPlanarFace;
                        }
                    }
                }

                // will come back later as needed.
                //
                //else if (geomObj is Instance)
                //{
                //}
                //else if (geomObj is Curve)
                //{
                //}
                //else if (geomObj is Mesh)
                //{
                //}
            }

            // if we come here, we did not find any.
            return null;
        }





    }
}
