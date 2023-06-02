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


            // (0) �t�@�~���e���v���[�g�̃`�F�b�N����
            if (!isRightTemplate(BuiltInCategory.OST_Columns))
            {
                Util.ErrorMsg("Please open ��(���[�g���P��).rft");
                return Result.Failed;
            }
            else
            {
                TaskDialog.Show("�t�@�~���e���v���[�g�̃`�F�b�N", "OK");
            }


            using (Transaction transaction = new Transaction(_rvtDoc))
            {
                try
                {
                    if (transaction.Start("CreateFamily") == TransactionStatus.Started)
                    {
                        // (1) ���o���̍쐬 
                        Extrusion pSolid = createSolid();

                        // We need to regenerate so that we can build on this new geometry
                        _rvtDoc.Regenerate();

                        // try this:
                        // if you comment addAlignment and addTypes calls below and execute only up to here,
                        // you will see the column's top will not follow the upper level.

                        // (2) �ʒu���킹�̒ǉ�
                        addAlignments(pSolid);

                        // try this: at each stage of adding a function here, you should be able to see the result in UI.

                        // (3) �^�C�v�̒ǉ�
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
            //���̃R�}���h�́A�t�@�~���[�G�f�B�^�[�̃R���e�L�X�g�ł̂݋@�\
            // �t�@�~���h�L�������g���ǂ������`�F�b�N
            if (!_rvtDoc.IsFamilyDocument)
            {
                Util.ErrorMsg("This command works only in the family editor.");
                return false;
            }

            // Check the template for an appropriate category here if needed.
            // �K�v�ł���΁A�����œK�؂ȃJ�e�S���̃e���v���[�g���m�F���Ă��������B
            // �e���v���[�g�����t�@�~����`�p���ǂ������`�F�b�N
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

            // �X�P�b�`����=��Ɩʂ̍쐬
            ReferencePlane pRefPlane = findElement (typeof(ReferencePlane), "�Q�Ɩ�") as ReferencePlane;
            SketchPlane pSketchPlane = SketchPlane.Create(_rvtDoc, pRefPlane.GetPlane());

            // ���o������
            double dHeight = mmToFeet(4000.0);

            // ���o��
            bool bIsSolid = true;  // ���true�B ���̃��\�b�h�̉��o���p�B
            return _rvtDoc.FamilyCreate.NewExtrusion(bIsSolid, pProfile, pSketchPlane, dHeight);
        }

        CurveArrArray createProfileRectangle()
        {
            double w = mmToFeet(600.0);
            double d = mmToFeet(600.0);
            const int nVerts = 4;   // �ӂ̐�

            // �l�p�`�̍쐬
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

            // �J�[�u����v���t�@�C�����쐬
            CurveArrArray pProfile = _rvtApp.Create.NewCurveArrArray();
            pProfile.Append(pLoop);

            return pProfile;
        }

        void addAlignments(Extrusion pBox)
        {
            // (1) ���̏�ʂ���Q�ƃ��x���ɍS��
            View pView = findElement(typeof(View), "����") as View;

            Level upperLevel = findElement(typeof(Level), "��Q�ƃ��x��") as Level;
            Reference ref1 = upperLevel.GetPlaneReference(); // ���x������Q�Ɩʂ��擾

            PlanarFace upperFace = findFace(pBox, new XYZ(0, 0, 1));
            Reference ref2 = upperFace.Reference;

            // �ʒu���킹�̍쐬
            _rvtDoc.FamilyCreate.NewAlignment(pView, ref1, ref2);

            // (2) ���[���x���������悤�ɍS��
            Level lowerLevel = findElement(typeof(Level),"���Q�ƃ��x��") as Level;
            Reference ref3 = lowerLevel.GetPlaneReference();

            PlanarFace lowerFace = findFace(pBox, new XYZ(0, 0, -1));
            Reference ref4 = lowerFace.Reference;

            // �ʒu���킹�̍쐬
            _rvtDoc.FamilyCreate.NewAlignment(pView, ref3, ref4);

            // (3) ���l�ɉE���O��̍S�����s��
            View pViewPlan = findElement(typeof(ViewPlan), "���Q�ƃ��x��") as View;

            // find reference planes
            ReferencePlane refRight = findElement(typeof(ReferencePlane),"�E") as ReferencePlane;
            ReferencePlane refLeft  = findElement(typeof(ReferencePlane), "��") as ReferencePlane;
            ReferencePlane refFront = findElement(typeof(ReferencePlane), "����") as ReferencePlane;
            ReferencePlane refBack  = findElement(typeof(ReferencePlane), "�w��") as ReferencePlane;

            // find the face of the box
            PlanarFace faceRight = findFace(pBox, new XYZ(1.0, 0.0, 0.0));
            PlanarFace faceLeft = findFace(pBox, new XYZ(-1.0, 0.0, 0.0));
            PlanarFace faceFront = findFace(pBox, new XYZ(0.0, -1.0, 0.0));
            PlanarFace faceBack = findFace(pBox, new XYZ(0.0, 1.0, 0.0));

            // �ʒu���킹�̍쐬
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

            // ���p�����[�^�̐ݒ�
            FamilyParameter paramW = pFamilyMgr.get_Parameter("��");
            double valW = mmToFeet(w);
            if (paramW != null) 
            {
                pFamilyMgr.Set(paramW, valW);
            }

            // ���s�p�����[�^�̐ݒ�
            FamilyParameter paramD = pFamilyMgr.get_Parameter("���s��");
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
