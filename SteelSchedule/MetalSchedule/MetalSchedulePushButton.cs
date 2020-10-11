#region Namespaces
using System;
using System.Collections.Generic;
using System.Collections;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Text;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Steel;
using Autodesk.Revit.DB.Structure;
using Autodesk.AdvanceSteel.DocumentManagement;
using Autodesk.AdvanceSteel.Geometry;
using Autodesk.AdvanceSteel.Modelling;
using Autodesk.AdvanceSteel.CADAccess;
using Autodesk.AdvanceSteel.Profiles;

using RVTDocument = Autodesk.Revit.DB.Document;
using ASDocument = Autodesk.AdvanceSteel.DocumentManagement.Document;
using RVTransaction = Autodesk.Revit.DB.Transaction;

using SteelSchedule.MetalSchedule;
#endregion

namespace SteelSchedule.MetalSchedule
{
    [Transaction(TransactionMode.Manual)]
    public class MetalSchedulePushButton : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            RVTDocument doc = commandData.Application.ActiveUIDocument.Document;
            Autodesk.Revit.DB.View calculateView = commandData.Application.ActiveUIDocument.ActiveView;

            List<SteelElement> steelelements = new List<SteelElement>();

            #region Проверка на наличие необходимых параметров

            bool paramOk = false;

            List<ProjectParameterData> result = new List<ProjectParameterData>();

            BindingMap map = doc.ParameterBindings;
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                ProjectParameterData newProjectParameterData
                  = new ProjectParameterData();

                if (it.Key.Name == "Масса погонного метра")
                {
                    newProjectParameterData.Definition = it.Key;
                    newProjectParameterData.Name = it.Key.Name;
                    newProjectParameterData.Binding = it.Current as ElementBinding;

                    result.Add(newProjectParameterData);
                }
            }

            foreach (ProjectParameterData p in result)
            {
                CategorySet cats = p.Binding.Categories;

                Category catCol = Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns);
                Category catBeam = Category.GetCategory(doc, BuiltInCategory.OST_StructuralFraming);

                TypeBinding typeBinding = p.Binding as TypeBinding;

                if (cats.Contains(catCol) & cats.Contains(catBeam) & typeBinding != null & p.Definition.ParameterType == ParameterType.MassPerUnitLength)
                {
                    paramOk = true;
                }
            }

            #endregion

            bool calcViewIs3D = (calculateView is Autodesk.Revit.DB.View3D);
            bool calcViewDetailLeveltree = false;
            try{ calcViewDetailLeveltree = (calculateView.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL).AsInteger() == 3); } catch { }

            if (calcViewIs3D && calcViewDetailLeveltree && paramOk == true)
            {
                try
                {
                    List<SteelProxyElement> rvtPlates = new FilteredElementCollector(doc, calculateView.Id)
                        .WhereElementIsNotElementType()
                        .OfClass(typeof(SteelProxyElement))
                        .Cast<SteelProxyElement>()
                        .Where(i => i.GeomType == GeomObjectType.Plate)
                        .ToList();

                    List<StructuralConnectionHandler> joints = new FilteredElementCollector(doc, calculateView.Id)
                        .WhereElementIsNotElementType()
                        .OfClass(typeof(StructuralConnectionHandler))
                        .Cast<StructuralConnectionHandler>()
                        .ToList();

                    List<BuiltInCategory> cat = new List<BuiltInCategory>();
                    cat.Add(BuiltInCategory.OST_StructuralColumns);
                    cat.Add(BuiltInCategory.OST_StructuralFraming);

                    ElementMulticategoryFilter multiFilter = new ElementMulticategoryFilter(cat);

                    List<Element> columnsBeams = new FilteredElementCollector(doc, calculateView.Id)
                        .WhereElementIsNotElementType()
                        .WherePasses(multiFilter)
                        .ToList();

                    #region Пластины из узлов
                    using (RvtDwgAddon.FabricationTransaction t = new RvtDwgAddon.FabricationTransaction(doc, true, "SteelSheetParam"))
                    {
                        foreach (StructuralConnectionHandler joint in joints)
                        {
                            if (joint.IsHidden(calculateView)) { continue; }

                            List<Subelement> subelems = joint.GetSubelements().ToList();

                            foreach (Subelement subelem in subelems)
                            {
                                Reference rf = subelem.GetReference();
                                FilerObject filerObj = this.GetFilerObject(doc, rf);

                                if (subelem.Category.Id == new ElementId(BuiltInCategory.OST_StructConnectionPlates))
                                {
                                    Plate pl = filerObj as Plate;

                                    ASPlate steelPlate = new ASPlate();
                                    steelPlate.Plate = pl;
                                    
                                    steelelements.Add(steelPlate);
                                }

                                if (subelem.Category.Id == new ElementId(BuiltInCategory.OST_StructConnectionProfiles))
                                {
                                    StraightBeam bm = filerObj as StraightBeam;

                                    ASProfile steelProfile = new ASProfile();
                                    steelProfile.StraightBeam = bm;
                                    steelelements.Add(steelProfile);
                                }

                            }
                        }
                    }
                    #endregion

                    #region Пластины, колонны, балки Revit
                    using (RVTransaction t = new RVTransaction(doc))
                    {
                        t.Start("SteelSheetParam");

                        foreach (SteelProxyElement plate in rvtPlates)
                        {
                            if (plate.IsHidden(calculateView)) { continue; }

                            Solid sol;

                            Options opt = new Options() { View = calculateView };
                            GeometryElement geoElem = plate.get_Geometry(opt);
                            sol = geoElem.First() as Solid;

                            ElementId mid = plate.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM).AsElementId();
                            Material material = doc.GetElement(mid) as Material;

                            RVTPlate rvtPlate = new RVTPlate();
                            rvtPlate.Plate = plate;
                            rvtPlate.Solid = sol;
                            rvtPlate.Material = material;
                            steelelements.Add(rvtPlate);
                        }

                        foreach (Element el in columnsBeams)
                        {
                            if (el.IsHidden(calculateView)) { continue; }

                            FamilyInstance inst = el as FamilyInstance;
                            Family fam = inst.Symbol.Family;

                            if (fam.get_Parameter(BuiltInParameter.FAMILY_STRUCT_MATERIAL_TYPE).AsInteger() == 1)
                            {
                                ElementType type = doc.GetElement(el.GetTypeId()) as ElementType;

                                ElementId materialId = el.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM).AsElementId();
                                ElementId typeMaterialId = type.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM).AsElementId();

                                Material material = doc.GetElement(materialId) as Material;

                                MainProfile mainProfile = new MainProfile();
                                mainProfile.Element = el;
                                mainProfile.Material = material;
                                mainProfile.TypeElement = type;

                                steelelements.Add(mainProfile);
                            }
                        }

                        t.Commit();
                    }
                    #endregion

                    steelelements = steelelements.ToList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                try
                {
                    DataRequestForm drf = new DataRequestForm(steelelements, commandData);
                    DialogResult dr = drf.ShowDialog();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                string eXmassage = "Ошибка: ";

                if(!calcViewIs3D)
                {
                    eXmassage = eXmassage + "\nОткройте 3Д вид! ";
                }

                if (!calcViewDetailLeveltree)
                {
                    eXmassage = eXmassage + "\nУстановите высокий уровень детализации! ";
                }

                if (paramOk == false)
                {
                    eXmassage = eXmassage + 
                        "\nОтсутствует параметр \"Масса погонного метра\". \nСоздайте парметр для типа категорий \"Каркас несущий\" и \"Несущие колонны\". Тип параметра \"Масса на единицу длинны\".";
                }

                MessageBox.Show(eXmassage);
            }

            return Result.Succeeded;
        }

        private FilerObject GetFilerObject(RVTDocument doc, Reference eRef)
        {
            FilerObject filerObject = null;
            ASDocument curDocAS = DocumentManager.GetCurrentDocument();
            if (null != curDocAS)
            {
                OpenDatabase currentDatabase = curDocAS.CurrentDatabase;
                if (null != currentDatabase)
                {
                    Guid uid = SteelElementProperties.GetFabricationUniqueID(doc, eRef);
                    string asHandle = currentDatabase.getUidDictionary().GetHandle(uid);
                    filerObject = FilerObject.GetFilerObjectByHandle(asHandle);
                }
            }
            return filerObject;
        }
    }
}
