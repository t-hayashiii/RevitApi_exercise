#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

#endregion

namespace IntroCs
{
    // DB Element - learn about Revit element
    [Transaction(TransactionMode.Manual)]
    public class DBElement : IExternalCommand
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

            // (1) pick an object on a screen. 画面上での選択
            Reference refPick = rvtUIDoc.Selection.PickObject(
                ObjectType.Element, "要素を選択してください");

            // we have picked somothing.
            Element elem = m_rvtDoc.GetElement(refPick);  // 選択オブジェクトの要素化

            // (2) let's see what kind of element we got.
            ShowBasicElementInfo(elem);

            // (3) identify each major types of element.
            IdentifyElement(elem);

            // (4) first parameters
            ShowParameters(elem, "要素のパラメータ");

            // check to see its type parameter as well タイプパラメータも
            ElementId elemTypeId = elem.GetTypeId();
            ElementType elemType = (ElementType)m_rvtDoc.GetElement(elemTypeId);
            ShowParameters(elemType, "タイプパラメータ");

            // accell to each parameters.
            RetrieveParameter(
                elem, "Element Parameter (by Name and BuiltInParameter");
            // the same logic applies to the type parameter.
            RetrieveParameter(
                elemType, "Type Parametere (by Name and BuiltInParameter");

            // (5) location
            ShowLocation(elem);

            // (6) geometry - the last piece. (Optional)
            ShowGeometry(elem); 

            return Result.Succeeded;
        }

        // 基本的な要素情報
        public void ShowBasicElementInfo(Element elem)
        {
            // let's see what kind of element we got.
            //
            string s = "選択要素:" + "\n";

            s += " クラス名 = " + elem.GetType().Name + "\n";     // クラス名
            s += " カテゴリ = " + elem.Category.Name + "\n";        // カテゴリ名
            s += " 要素ID = " + elem.Id.ToString() + "\n" + "\n"; // ID

            // and, check its type info.
            //
            // Dim elemType As ElementType = elem.ObjectType '' this is obsolete.
            ElementId elemTypeId = elem.GetTypeId();
            ElementType elemType = (ElementType)m_rvtDoc.GetElement(elemTypeId);

            s += "選択要素のタイプ:" + "\n";
            s += " クラス名 = " + elemType.GetType().Name + "\n";
            s += " カテゴリ = " + elemType.Category.Name + "\n";
            s += " 要素タイプID = " + elemType.Id.ToString() + "\n";

            // finally show it.

            TaskDialog.Show("Basic Element Info", s);
        }

        // 要素の識別_システムファミリとそれ以外の違い
        public void IdentifyElement (Element elem)
        {
            string s = ""; 

            if (elem is Wall)
            {
                s = "Wall";
            }
            else if (elem is Floor)
            {
                s = "Floor";
            }
            else if (elem is RoofBase)
            {
                s = "Roof";
            }
            else if (elem is FamilyInstance)
            {
                if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
                {
                    s = "Door";
                }
                else if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
                {
                    s = "Window";
                }
                else if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Furniture)
                {
                    s = "Furniture";
                }
                else
                {
                    s = "Component family instance";
                }
            }
            
            else if (elem is HostObject)
            {
                s = "System family instance";
            }
            else
            {
                s = "other";
            }

            s = "選択した要素: " + s;

            // show it.
            TaskDialog.Show("Identify Element", s);
        }
    
        // 要素のパラメータ値一覧を表示
        public void ShowParameters(Element elem, string header)
        {
            IList<Parameter> paramSet = elem.GetOrderedParameters();  // プロパティで表示されるパラメータのコレクション
            string s = string.Empty;  // パラメータ値を入れる変数を準備

            foreach (Parameter param in paramSet)
            {
                string name = param.Definition.Name;  // パラメータ名
                // 下記に定義したメソッドを参照
                string val = ParameterToString(param);
                s += name + " = " + val + "\n";
            }
            TaskDialog.Show(header, s);
        }
        
        // ParameterToStringメソッドの定義
        public static string ParameterToString(Parameter param)
        {
            string val = "none";  // 空の際に表示する文字型の文字
            if (param == null)
            {
                return val; 
            }
            // パラメータ値を取得するために型によって場合分け
            switch (param.StorageType)
            {
                case StorageType.Double:
                    double dVal = param.AsDouble();
                    val = dVal.ToString();
                    break;

                case StorageType.Integer:
                    int iVal = param.AsInteger();   
                    val = iVal.ToString();
                    break;

                case StorageType.String:
                    string sVal = param.AsString();
                    val = sVal;
                    break;

                case StorageType.ElementId:
                    ElementId idVal = param.AsElementId();
                    val = idVal.IntegerValue.ToString();
                    break;

                default:
                    break;
            }
            return val;
        }
        
        // 特定のパラメータ値の抽出
        public void RetrieveParameter(Element elem, string header)
        {
            string s = string.Empty;

            // (1) by BuiltInParameter
            Parameter param =
                elem.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
            if (param != null)
            {
                s += "コメント (by BuiltInParameter) = " + ParameterToString(param) + "\n";
            }

            // (2) by name.(Mark - most of instance has this parameter.)
            // if you use this method, it will language specific.
            param = elem.LookupParameter("マーク");  
            if (param != null)
            {
                s += "マーク (by Name) = " + ParameterToString(param) + "\n";
            }

            // the following should be in most of type parameter
            param = elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS);
            if (param != null)
            {
                s += "コメント(タイプ) (by BuiltInParameter) = " + ParameterToString(param) + "\n";
            }

            param = elem.LookupParameter("耐火等級");
            if (param != null) 
            {
                s += "耐火等級 (by Name) = " + ParameterToString(param) + "\n";
            }

            // using the BuiltInParameter, you can sometimes access one that is 
            // not in the parameters set.
            // Note : this works only for element type.
            param = elem.get_Parameter(
                BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM);
            if (param != null) 
            {
                s += "ファミリとタイプ名 (only by BuiltInParameter) = " +
                    ParameterToString(param) + "\n";
            }

            // show it.
            TaskDialog.Show(header, s);
        }

        // 位置情報の抽出
        public void ShowLocation(Element elem) 
        {
            string s = "位置情報: " + "\n" + "\n";
            Location loc = elem.Location;

            // 位置情報が点の場合
            if (loc is LocationPoint)
            {
                LocationPoint locPoint = (LocationPoint)loc;
                XYZ pt = locPoint.Point;        // ジオメトリ座標
                double r = locPoint.Rotation;   // 回転数値

                s += "点の位置" + "\n";
                s += "Point = " + PointToString(pt) + "\n";
                s += "Rotation = " + r.ToString() + "\n";
            }
            else if (loc is LocationCurve)
            {
                LocationCurve locCurve = (LocationCurve)loc;
                Curve crv = locCurve.Curve;     // ジオメトリカーブ


                s += "線の位置" + "\n";
                s += "EndPoint(0)/始点 = " +
                    PointToString(crv.GetEndPoint(0)) + "\n";       // 始点
                s += "EndPoint(1)/終点 = " +
                    PointToString(crv.GetEndPoint(1)) + "\n";       // 終点
                s += "線の長さ = " + crv.Length.ToString() + "\n";    // 長さ

                // Location Curve also has property JoinType at the end
                s += "JoinType(0) = " + locCurve.get_JoinType(0).ToString() + "\n";
                s += "JoinType(1) = " + locCurve.get_JoinType(1).ToString() + "\n";
            }

            // show it.
            TaskDialog.Show("位置の表示", s);
        }

        // PointToStringメソッドの定義
        public static string PointToString(XYZ pt)
        {

            return pt.ToString();
        }

        // 要素のジオメトリを抽出
        public void ShowGeometry(Element elem)
        {
            // Set a geometry option
            Options opt = m_rvtApp.Create.NewGeometryOptions();
            opt.DetailLevel = ViewDetailLevel.Fine;

            // Get the geometry from element
            GeometryElement geoElem = elem.get_Geometry(opt);

            // if there is a geometry data, retrieve it as a string to show it.
            string s = (geoElem == null) ?
                "no data" :
                GeometryElementToString(geoElem);

            TaskDialog.Show("要素のジオメトリ", s);
        }

        // GeometryElementToString メソッドの定義
        public static string GeometryElementToString(GeometryElement geomElem)
        {
            string str = string.Empty;

            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid)
                {
                    str += "Solid" + "\n";
                }

                else if (geomObj is GeometryInstance)
                {
                    // 例: ドア、窓                    
                    str += " -- Geometry.Instance -- " + "\n";
                    GeometryInstance geomInstance = (GeometryInstance)geomObj;  // ?
                    GeometryElement geoElem = geomInstance.SymbolGeometry;      // ?

                    str += GeometryElementToString(geoElem);    // ?
                }
                
                else if (geomObj is Curve)
                {
                    str += "Curve" + "\n";
                }

                else if (geomObj is Mesh)
                {
                    str += "Mesh" + "\n";
                }

                else
                {
                    str += " *** unknown geometry type" + geomObj.GetType().Name;
                }
            }

            return str;
        }
    }
}
