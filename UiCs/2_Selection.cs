#region Namespaces
using System;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;  // 選択処理用
// using Util;
using System.Collections;
#endregion

namespace UiCs
{
    [Transaction(TransactionMode.Manual)]
    /// <summary>Main
    /// </summary>
    public class UISelection : IExternalCommand
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

            /// 現在選択中の要素を表示・取得
            /// 
            ICollection<ElementId> selectedElementIds = _uiDoc.Selection.GetElementIds();
            TaskDialog.Show("Revit", "選択中の要素の数 : " + selectedElementIds.Count.ToString());
            ShowElementList(selectedElementIds, "選択済要素");

            try
            {
                // (2.1) pick methods basics. 
                // there are four types of pick methods: PickObject, PickObjects, PickElementByRectangle, PickPoint. 
                // Let's quickly try them out. 

                PickMethodsBasics();

                // (2.2) selection object type 
                // in addition to selecting objects of type Element, the user can pick faces, edges, and point on element. 

                // PickFaceEdgePoint();

                // (2.3) selection filter 
                // if you want additional selection criteria, such as only to pick a wall, you can use selection filter. 

                // ApplySelectionFilter();
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                TaskDialog.Show("UI selection", "You have canceled selection.");
            }
            catch (Exception)
            {
                TaskDialog.Show("UI selection", "Some other exception caught in CancelSelection()");
            }

            // (2.4) canceling selection 
            // when the user cancel or press [Esc] key during the selection, OperationCanceledException will be thrown. 

            CancelSelection();

            // (3) apply what we learned to our small house creation 
            // we put it as a separate command. See at the bottom of the code. 
            // CreateHouseUI



            return Result.Succeeded;
        }





        


        /// <summary>
        /// Pick methods sampler. 
        /// Quickly try: PickObject, PickObjects, PickElementByRectangle, PickPoint. 
        /// Without specifics about objects we want to pick. 
        /// </summary>
        public void PickMethodsBasics()
        {
            // (1) Pick Object (we have done this already. But just for the sake of completeness.) 
            PickMethod_PickObject();    

            // (2) Pick Objects 
            PickMethod_PickObjects();

            // (3) Pick Element By Rectangle 
            PickMethod_PickElementByRectangle();

            // (4) Pick Point 
            PickMethod_PickPoint();

            PickPointOnElement();

            //
            Pickface();     // 面を選択
            PickEdge();     // エッジを選択

            //
            PickWall();     // 壁を選択(制限付き)


        }


        /// PickMethod_PickObject()
        /// 
        public void PickMethod_PickObject()
        {
            Reference r = _uiDoc.Selection.PickObject(
                ObjectType.Element, "要素を選択してください");

            Element e = _uiDoc.Document.GetElement(r);

            ShowBasicElementInfo(e, "選択した要素:");
        }

        /// PickMethod_PickObjects()
        /// 
        public void PickMethod_PickObjects()
        {
            IList<Reference> refs = _uiDoc.Selection.PickObjects(
                ObjectType.Element, "複数の要素を選択してください");

            // Put it in a List form.
            IList<Element> elems = new List<Element>();
            foreach (Reference r in refs) 
            {
                elems.Add(_uiDoc.Document.GetElement(r));
            }

            ShowElementList_t(elems, "選択した要素群");  // メソッドの引数は要素 not ID)
        }

        /// PickMethod_PickElementByRectangle()
        /// 
        public void PickMethod_PickElementByRectangle()
        {
            IList<Element> elems = _uiDoc.Selection.PickElementsByRectangle("矩形で選択してください");

            ShowElementList_t(elems, "矩形で選択した要素:");
        }

        /// PickMethond_PickPoint()
        /// 
        public void PickMethod_PickPoint()
        {
            XYZ pt = _uiDoc.Selection.PickPoint("点を選択してください");

            string msg = "選択した点: ";
            msg += PointToString(pt);

            TaskDialog.Show("PickPoint", msg);
        }

        /// PickPointOnElement
        /// 
        public void PickPointOnElement()
        {
            Reference r = _uiDoc.Selection.PickObject(
                ObjectType.PointOnElement, "要素上の点を選択してください");

            Element e = _uiDoc.Document.GetElement(r);
            XYZ pt = r.GlobalPoint;     // リファレンスがヒットする位置

            string msg = "";
            if(pt != null)
            {
                msg = "You picked the point" + PointToString(pt) 
                    + " on an element " + e.Id.ToString() + "\r\n" ;
            }
            else
            {
                msg = "no Point picked \n";
            }

            TaskDialog.Show("PickPointOnElement", msg);
        }

        /// PickFace()
        /// 
        public void Pickface()
        {
            Reference r = _uiDoc.Selection.PickObject(
                ObjectType.Face, "面を選択してください");
            Element e = _uiDoc.Document.GetElement(r);

            Face oFace = e.GetGeometryObjectFromReference(r) as Face;

            string msg = "";
            if(oFace != null)
            {
                msg = "You picked the face of element " + e.Id.ToString() + "\r\n" ;  
            }
            else
            {
                msg = "no Face picked \n";
            }

            TaskDialog.Show("PickFace", msg);
        }

        /// PickEdge()
        /// 
        public void PickEdge()
        {
            Reference r = _uiDoc.Selection.PickObject(ObjectType.Edge, "エッジを選択してください");
            Element e = _uiDoc.Document.GetElement(r);
            //Edge oEdge = r.GeometryObject as Edge; // 2011
            Face oEdge = e.GetGeometryObjectFromReference(r) as Face; // 2012

            // Show it. 
            string msg = "";
            if (oEdge != null)
            {
                msg = "You picked an edge of element " + e.Id.ToString() + "\r\n";
            }
            else
            {
                msg = "no Edge picked \n";
            }

            TaskDialog.Show("PickEdge", msg);
        }


        /// PickWall()
        /// 
        public void PickWall()
        {
            SelectionFilterWall selFilterWall = new SelectionFilterWall();
            Reference r = _uiDoc.Selection.PickObject(ObjectType.Element,
                selFilterWall, "壁を選択してください");

            Element e = _uiDoc.Document.GetElement(r);

            ShowBasicElementInfo(e);
        }


        /// <summary>
        /// Canceling selection 
        /// When the user presses [Esc] key during the selection, OperationCanceledException will be thrown. 
        /// </summary>
        public void CancelSelection()
        {
            try
            {
                Reference r = _uiDoc.Selection.PickObject(ObjectType.Element, "Select an element, or press [Esc] to cancel");
                Element e = _uiDoc.Document.GetElement(r);

                ShowBasicElementInfo(e);
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                TaskDialog.Show("CancelSelection", "選択をキャンセルしました");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("CancelSelection", "Other exception caught in CancelSelection(): " + ex.Message);
            }
        }



        #region "Helper Function"

        /// <summary>
        /// Show basic information about the given element. 
        /// </summary>
        public void ShowBasicElementInfo(Element e)
        {
            // Let's see what kind of element we got. 
            string s = "You picked: \n";

            s += ElementToString(e);

            // Show what we got. 

            TaskDialog.Show("Revit UI Lab", s);
        }

        /// ShowBasicElementInfo 
        ///
        public void ShowBasicElementInfo(Element e, string header)
        {
            // Let's see what kind of element we got. 
            string s = "\n\n - Class - Category - Name (or Family: Type Name) - Id - " + "\r\n";
            s += ElementToString(e);

            s = header + s;

            // Show what we got. 
            TaskDialog.Show("Revit UI Lab", s);
        }

        /// ShowElementList
        /// 
        public void ShowElementList(IEnumerable elemIds, string header)
        {
            string s = "\n\n - Class - Category - Name (or Family: Type Name) - Id - " + "\r\n";

            int count = 0;
            foreach (ElementId eId in elemIds)
            {
                count++;
                Element e = _uiDoc.Document.GetElement(eId);
                s += ElementToString(e);
            }

            s = header + "(" + count.ToString() + ")" + s;

            TaskDialog.Show("Revit UI Lab", s);
        }

        /// ShowElementList_t 要素から
        /// 
        public void ShowElementList_t(IEnumerable elems, string header)
        {
            string s = "\n\n - Class - Category - Name (or Family: Type Name) - Id - " + "\r\n";

            int count = 0;
            foreach (Element e in elems)
            {
                count++;
                s += ElementToString(e);
            }

            s = header + "(" + count.ToString() + ")" + s;

            TaskDialog.Show("Revit UI Lab", s);
        }


        /// ElementToString
        /// 
        public string ElementToString(Element e)
        {
            if (e == null)
            {
                return "none";
            }

            string name = "";

            if (e is ElementType)   // 要素がタイプなら
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

            return e.GetType().Name + "; " + e.Category.Name + "; " + name + "; " + e.Id.IntegerValue.ToString() + "\r\n";
        }

        /// PointToString
        /// 
        public static string PointToString(XYZ pt)
        {
            if (pt == null)
            {
                return "";
            }

            return "(" + pt.X.ToString("F2") + ", " + pt.Y.ToString("F2") + ", " + pt.Z.ToString("F2") + ")";
        }

        #endregion
    }



    /// Class SelectionFilterWall
    /// 
    class SelectionFilterWall : ISelectionFilter
    {
        public bool AllowElement(Element e)
        {
            return e is Wall;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return true;
        }
    }




}
