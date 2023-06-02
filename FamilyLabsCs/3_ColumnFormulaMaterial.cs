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
    class RvtCmd_FamilyCreateColumnFormulaMaterial : IExternalCommand
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
                        // (1.1) add reference planes
                        addReferencePlanes();

                        // (1.2) create a simple extrusion. This time we create a L-shape.
                        Extrusion pSolid = createSolid();
                        _rvtDoc.Regenerate();

                        // (2) add alignment
                        addAlignments(pSolid);

                        // (3.1) add parameters
                        addParameters();

                        // (3.2) add dimensions
                        addDimensions();

                        // (3.3) add types
                        addTypes();

                        // (4.1) add formula
                        addFormulas();

                        // (4.2) add materials
                        addMaterials(pSolid);


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

        void addReferencePlanes()
        {
            double tw = mmToFeet(150);      // 幅方向のオフセット
            double td = mmToFeet(150);      // 奥行き方向のオフセット

            // (1) 水平方向の参照面
            View pViewPlan = findElement (typeof (ViewPlan), "下参照レベル") as View;
            ReferencePlane refFront = findElement(typeof(ReferencePlane), "正面") as ReferencePlane;

            //  get the bubble and free ends from front ref plan and offset by td.
            XYZ p1 = refFront.BubbleEnd;
            XYZ p2 = refFront.FreeEnd;
            XYZ pBubbleEnd = new XYZ(p1.X, p1.Y + td, p1.Z);
            XYZ pFreeEnd = new XYZ(p2.X, p2.Y + td, p2.Z);

            // "OffsetH"という参照面を作成
            ReferencePlane refPlane =
                _rvtDoc.FamilyCreate.NewReferencePlane(
                    pBubbleEnd, pFreeEnd, XYZ.BasisZ, pViewPlan);
            refPlane.Name = "OffsetH";

            // (2) 垂直方向の参照面
            ReferencePlane refLeft = findElement(typeof(ReferencePlane), "左") as ReferencePlane;

            //  get the bubble and free ends from front ref plan and offset by tw.
            p1 = refLeft.BubbleEnd;
            p2 = refLeft.FreeEnd;
            pBubbleEnd = new XYZ(p1.X + tw, p1.Y, p1.Z);
            pFreeEnd   = new XYZ(p2.X + tw, p2.Y, p2.Z);

            // "OffsetV"という参照面を作成
            refPlane =
                _rvtDoc.FamilyCreate.NewReferencePlane(
                    pBubbleEnd, pFreeEnd, XYZ.BasisZ, pViewPlan);
            refPlane.Name = "OffsetV";

        }

        Extrusion createSolid()
        {
            CurveArrArray pProfile = createProfileLShape();

            // スケッチ平面=作業面の作成
            ReferencePlane pRefPlane = findElement (typeof(ReferencePlane), "参照面") as ReferencePlane;
            SketchPlane pSketchPlane = SketchPlane.Create(_rvtDoc, pRefPlane.GetPlane());

            // 押出し高さ
            double dHeight = mmToFeet(4000.0);

            // 押出し
            bool bIsSolid = true;  // 常にtrue。 このメソッドの押出し用。
            return _rvtDoc.FamilyCreate.NewExtrusion(bIsSolid, pProfile, pSketchPlane, dHeight);
        }

        CurveArrArray createProfileLShape()
        {
            double w = mmToFeet(600.0);
            double d = mmToFeet(600.0);
            double tw = mmToFeet(150);
            double td = mmToFeet(150);
            const int nVerts = 6;   // 辺の数

            // 四角形の作成
            XYZ[] pts = new XYZ[]
            {
                new XYZ(-w/2.0, -d/2.0, 0.0),
                new XYZ(w/2.0, -d/2.0, 0.0),
                new XYZ(w/2.0, -d/2.0 + td, 0.0),
                new XYZ(-w/2.0 + tw, -d/2.0 + td, 0.0),
                new XYZ(-w/2.0 + tw, d/2.0, 0.0),
                new XYZ(-w/2.0 , d/2.0, 0.0),
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
            ReferencePlane refOffsetV = findElement(typeof(ReferencePlane), "OffsetV") as ReferencePlane;
            ReferencePlane refOffsetH = findElement(typeof(ReferencePlane), "OffsetH") as ReferencePlane;

            // find the face of the box ボックスの面
            PlanarFace faceRight = findFace(pBox, new XYZ(1.0, 0.0, 0.0), refRight);
            PlanarFace faceLeft = findFace(pBox, new XYZ(-1.0, 0.0, 0.0));
            PlanarFace faceFront = findFace(pBox, new XYZ(0.0, -1.0, 0.0));
            PlanarFace faceBack = findFace(pBox, new XYZ(0.0, 1.0, 0.0), refBack);
            PlanarFace faceOffsetV = findFace(pBox, new XYZ(1.0, 0.0, 0.0), refOffsetV);
            PlanarFace faceOffsetH = findFace(pBox, new XYZ(0.0, 1.0, 0.0), refOffsetH);

            // 位置合わせの作成
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refRight.GetReference(), faceRight.Reference);
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refLeft.GetReference(), faceLeft.Reference);
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refFront.GetReference(), faceFront.Reference);
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refBack.GetReference(), faceBack.Reference);
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refOffsetV.GetReference(), faceOffsetV.Reference);
            _rvtDoc.FamilyCreate.NewAlignment(pViewPlan, refOffsetH.GetReference(), faceOffsetH.Reference);

        }

        void addParameters()
        {
            FamilyParameter paramTw = _rvtDoc.FamilyManager.AddParameter(
                "Tw", BuiltInParameterGroup.PG_GEOMETRY, ParameterType.Length, false);
            FamilyParameter paramTd = _rvtDoc.FamilyManager.AddParameter(
                "Td", BuiltInParameterGroup.PG_GEOMETRY, ParameterType.Length, false);

            // 初期値の設定
            double tw = mmToFeet(150);
            double td = mmToFeet(150);
            _rvtDoc.FamilyManager.Set(paramTw, tw);
            _rvtDoc.FamilyManager.Set(paramTd, td);
        }

        void addDimensions()
        {
            // find the plan view
            //
            View pViewPlan = findElement(typeof(ViewPlan), "下参照レベル") as View;

            // find reference planes
            //
            ReferencePlane refLeft    = findElement(typeof(ReferencePlane), "左") as ReferencePlane;
            ReferencePlane refFront   = findElement(typeof(ReferencePlane), "正面") as ReferencePlane;
            ReferencePlane refOffsetV = findElement(typeof(ReferencePlane), "OffsetV") as ReferencePlane;
            ReferencePlane refOffsetH = findElement(typeof(ReferencePlane), "OffsetH") as ReferencePlane;

            // (1) 寸法の作成 1
            // 

            // 寸法位置の線
            //
            XYZ p0 = refLeft.FreeEnd;  // 基準面の自由端
            XYZ p1 = refOffsetV.FreeEnd;
            Line pLine = Line.CreateBound(p0, p1);

            // 参照
            //
            ReferenceArray pRefArray = new ReferenceArray();  // 参照面配列の定義
            pRefArray.Append(refLeft.GetReference());
            pRefArray.Append(refOffsetV.GetReference());    

            // 寸法の作成
            //
            Dimension pDimTw = _rvtDoc.FamilyCreate.NewDimension(pViewPlan, pLine, pRefArray);

            // 寸法にラベル(パラメータ)を追加
            //
            FamilyParameter paramTw = _rvtDoc.FamilyManager.get_Parameter("Tw");
            pDimTw.FamilyLabel = paramTw;

            // -------------------------------------------------------------------
            // (2) 寸法の作成 2
            // 

            // 寸法位置の線
            //
            p0 = refFront.FreeEnd;  // 基準面の自由端
            p1 = refOffsetH.FreeEnd;
            pLine = Line.CreateBound(p0, p1);

            // 参照
            //
            pRefArray = new ReferenceArray();  // 参照面配列の定義
            pRefArray.Append(refFront.GetReference());
            pRefArray.Append(refOffsetH.GetReference());

            // 寸法の作成
            //
            Dimension pDimTd = _rvtDoc.FamilyCreate.NewDimension(pViewPlan, pLine, pRefArray);

            // 寸法にラベル(パラメータ)を追加
            //
            FamilyParameter paramTd = _rvtDoc.FamilyManager.get_Parameter("Td");
            pDimTd.FamilyLabel = paramTd;
        }

        void addTypes()
        {
            addType("600x900", 600, 900, 150, 225);
            addType("1000x300", 1000, 300, 250, 75);
            addType("600x600", 600, 600, 150, 150);
        }

        void addType(string name, double w, double d, double tw, double td) 
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

            // Twパラメータの設定
            FamilyParameter paramTw = pFamilyMgr.get_Parameter("Tw");
            double valTw = mmToFeet(tw);
            if (paramTw != null)
            {
                pFamilyMgr.Set(paramTw, valTw);
            }

            // Tdパラメータの設定
            FamilyParameter paramTd = pFamilyMgr.get_Parameter("Td");
            double valTd = mmToFeet(td);
            if (paramTd != null)
            {
                pFamilyMgr.Set(paramTd, valTd);
            }
        }

        public void addFormulas()
        {
            FamilyManager pFamilyMgr = _rvtDoc.FamilyManager;

            // get the parameter
            FamilyParameter paramTw = pFamilyMgr.get_Parameter("Tw");
            FamilyParameter paramTd = pFamilyMgr.get_Parameter("Td");

            // set the formula
            pFamilyMgr.SetFormula(paramTw, "幅 / 4.0");
            pFamilyMgr.SetFormula(paramTd, "奥行き / 4.0");
        }
        
        public void addMaterials(Extrusion pSolid)
        {
            // (1) get the materials id that we are interested in (例 ガラス)
            Material pMat = findElement(typeof (Material), "ガラス") as Material;  
            if (pMat != null)
            {
                ElementId idMat = pMat.Id;

                // ファミリのパラメータ追加
                FamilyManager pFamilyMgr = _rvtDoc.FamilyManager;
                FamilyParameter famParamFinish =
                    pFamilyMgr.AddParameter("柱仕上", BuiltInParameterGroup.PG_MATERIALS,
                    ParameterType.Material, true);

                // パラメータの関連付け
                Parameter paramMat = pSolid.LookupParameter("マテリアル");
                pFamilyMgr.AssociateElementParameterToFamilyParameter(paramMat,
                    famParamFinish);

                // タイプの追加(ガラス)
                addType("ガラス", 600, 600, 100, 100);
                pFamilyMgr.Set(famParamFinish, idMat);

            }

        }



        /// <summary>
        /// ヘルパー関数
        /// </summary>

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

        PlanarFace findFace(Extrusion pBox, XYZ normal, ReferencePlane refPlane)  // 追加関数。引数3つ。
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
                if (geomObj is Solid)  // solid is what we are interested in.
                {
                    Solid pSolid = geomObj as Solid;
                    FaceArray faces = pSolid.Faces;
                    foreach (Face pFace in faces)
                    {
                        PlanarFace pPlanarFace = (PlanarFace)pFace;
                        // check to see if they have same normal
                        if ((pPlanarFace != null) && pPlanarFace.FaceNormal.IsAlmostEqualTo(normal))
                        {
                            // additionally, we want to check if the face is on the reference plane
                            //
                            XYZ p0 = refPlane.BubbleEnd;
                            XYZ p1 = refPlane.FreeEnd;
                            //Line pCurve = _app.Create.NewLineBound(p0, p1);  // Revit 2013
                            Line pCurve = Line.CreateBound(p0, p1);  // Revit 2014
                            if (pPlanarFace.Intersect(pCurve) == SetComparisonResult.Subset)
                            {
                                return pPlanarFace; // we found the face
                            }
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
