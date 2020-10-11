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
#endregion

namespace SteelSchedule.MetalSchedule
{
    public abstract class SteelElement
    {
        public abstract string GetGost();
        public abstract string GetMetal();
        public abstract string GetName();
        public abstract double GetMass();
        public abstract string GetUses();

        public string Gost
        {
            get { return this.GetGost(); }
        }

        public string Metal
        {
            get { return this.GetMetal(); }
        }

        public string Name
        {
            get { return this.GetName(); }
        }

        public double Mass
        {
            get { return this.GetMass(); }
        }

        public string Uses
        {
            get { return this.GetUses(); }
        }

        public double density = 7850;

        protected Dictionary<string, string> GostNames
        {
            get
            {
                Dictionary<string, string> _gostNames = new Dictionary<string, string>();

                _gostNames.Add("W", "Уголки стальные горячекатанные неравнополочные ГОСТ 8510-93");
                _gostNames.Add("Ws", "Уголки стальные горячекатанные равнополочные ГОСТ 8509-93");
                _gostNames.Add("U", "Швеллеры стальные гнутые равнополочные ГОСТ 8278-83");
                _gostNames.Add("H", "Профили стальные гнутые замкнутые сварные квадратные и прямоугольные ГОСТ 30245-2003");
                _gostNames.Add("t", "Прокат листовой горячекатанный ГОСТ 19903-2015");
                _gostNames.Add("I", "Двутавр");

                return _gostNames;
            }
        }
    }

    public class MainProfile : SteelElement
    {
        private Element _element;
        private Material _material;
        private ElementType _typeElem;

        private string _gost;
        private string _metal;
        private string _name;
        private double _mass;
        private string _uses;

        private double _cutlenght;
        private double _desSecArea;
        private double _volume;

        public override string GetGost()
        {
            _gost = _typeElem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString();
            if (_gost == null) { _gost = "Комментарии к типоразмеру"; }
            return _gost;
        }

        public override string GetMetal()
        {
            if (this.Material == null)
            {
                _metal = "Не задан материал.";
            }
            else
            {
                _metal = this.Material.Name;
                _metal = Regex.Replace(_metal.Split()[0], @"[^0-9a-zA-Z0-9А-Яа-я-\ ]+", " ") + "\nГОСТ27772";
            }

            return _metal;
        }

        public override string GetName()
        {
            _name = this.Element.Name;
            return _name;
        }

        public override double GetMass()
        {
                double kiloPerMeter = _typeElem.LookupParameter("Масса погонного метра").AsDouble() * 3.280839895013;

                _mass = kiloPerMeter * this.CutLenght;
                return _mass;
        }

        public override string GetUses()
        {
            if (this.Material == null)
            {
                _uses = "Не задан материал.";
            }
            else
            {
                _uses = this.Material.Name;
                _uses = _uses.Remove(0, _uses.IndexOf(' ') + 1);
            }

            return _uses;
        }

        public Element Element
        {
            get { return _element; }
            set { _element = value; }
        }

        public Material Material
        {
            get { return _material; }
            set { _material = value; }
        }

        public ElementType TypeElement
        {
            get { return _typeElem; }
            set { _typeElem = value; }
        }

        private double Volume
        {
            get
            {
                if (this.Element.Category.Name == "Каркас несущий")
                {
                    double volCacl = UnitUtils.Convert(this.Element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble(), DisplayUnitType.DUT_CUBIC_FEET, DisplayUnitType.DUT_CUBIC_METERS);
                    volCacl = Math.Round(volCacl, 15);

                    double volFact = this.CutLenght * this.DesSecArea;
                    volFact = Math.Round(volFact, 15);

                    if (volCacl >= volFact)
                    {
                        _volume = volCacl;
                    }
                    else
                    {
                        _volume = volFact;
                    }
                }

                if (this.Element.Category.Name == "Несущие колонны")
                {
                    _volume = UnitUtils.Convert(this.Element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble(), DisplayUnitType.DUT_CUBIC_FEET, DisplayUnitType.DUT_CUBIC_METERS);
                    _volume = Math.Round(_volume, 15);
                }

                return _volume;
            }
        }

        private double CutLenght
        {
            get
            {
                if (this.Element.Category.Name == "Каркас несущий")
                {
                    _cutlenght = UnitUtils.Convert(this.Element.get_Parameter(BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH).AsDouble(), DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES, DisplayUnitType.DUT_METERS);
                }

                if (this.Element.Category.Name == "Несущие колонны")
                {
                    _cutlenght = UnitUtils.Convert(this.Element.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsDouble(), DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES, DisplayUnitType.DUT_METERS);
                }

                return _cutlenght;
            }
        }

        private double DesSecArea
        {
            get
            {
                _desSecArea = UnitUtils.Convert(this.TypeElement.LookupParameter("A").AsDouble(), DisplayUnitType.DUT_SQUARE_FEET, DisplayUnitType.DUT_SQUARE_METERS);
                return _desSecArea;
            }
        }
    }

    public class ASProfile : SteelElement
    {
        private StraightBeam _straightBeam;

        private string _gost;
        private string _metal;
        private string _name;
        private double _mass;
        private string _uses;

        private double _volume;

        public override string GetGost()
        {
            GostNames.TryGetValue(this.StraightBeam.GetProfType().Type.ToString(), out _gost);
            if (_gost == null) { _gost = " "; }
            return _gost;
        }

        public override string GetMetal()
        {
            _metal = this.StraightBeam.MaterialDescription;
            _metal = Regex.Replace(_metal.Split()[0], @"[^0-9a-zA-Z0-9А-Яа-я-\ ]+", " ") + "\nГОСТ27772";
            return _metal;
        }

        public override string GetName()
        {
            int i = 1;
            _name = this.StraightBeam.GetName(out i).ToString();
            return _name;
        }

        public override double GetMass()
        {
            _mass = this.Volume * density;
            return _mass;
        }

        public override string GetUses()
        {
            _uses = this.StraightBeam.MaterialDescription;
            _uses = _uses.Remove(0, _uses.IndexOf(' ') + 1);
            return _uses;
        }

        public StraightBeam StraightBeam
        {
            get { return _straightBeam; }
            set { _straightBeam = value; }
        }

        private double Volume
        {
            get
            {
                ProfileSectionValues psv = new ProfileSectionValues();
                psv = this.StraightBeam.GetProfType().GetStructuralData();

                _volume = psv.A * this.StraightBeam.SysLength;
                _volume = UnitUtils.Convert(_volume, DisplayUnitType.DUT_CUBIC_MILLIMETERS, DisplayUnitType.DUT_CUBIC_METERS);

                _volume = Math.Round(_volume, 15);
                return _volume;
            }
        }
    }

    public class RVTPlate : SteelElement
    {
        private Solid _solid;
        private SteelProxyElement _plate;
        private Material _material;

        private string _gost;
        private string _metal;
        private string _name;
        private double _mass;
        private string _uses;

        private double _volume;

        public override string GetGost()
        {
            GostNames.TryGetValue("t", out _gost);
            return _gost;
        }

        public override string GetMetal()
        {
            _metal = Regex.Replace(this.Material.Name.Split()[0], @"[^0-9a-zA-Z0-9А-Яа-я-\ ]+", " ") + "\nГОСТ27772";
            return _metal;
        }

        public override string GetName()
        {
            BoundingBoxXYZ bb = OOBB.GetOOBB(this.Solid);
            XYZ sizes = GetSizes(bb);

            double z = Math.Round(UnitUtils.Convert(sizes.Z, DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES, DisplayUnitType.DUT_MILLIMETERS));

            if (z >= 25)
            {
                _name = "t" + z + " z25";
            }
            else
            {
                _name = "t" + z;
            }
            return _name;
        }

        public override double GetMass()
        {
            _mass = this.Volume * density;
            return _mass;
        }

        public override string GetUses()
        {
            _uses = this.Material.Name;
            _uses = _uses.Remove(0, _uses.IndexOf(' ') + 1);
            return _uses;
        }

        public SteelProxyElement Plate
        {
            get { return _plate; }
            set { _plate = value; }
        }

        public Material Material
        {
            get { return _material; }
            set { _material = value; }
        }
        	
        public Solid Solid
        {
            get { return _solid; }
            set { _solid = value; }
        }

        private double Volume
        {
            get
            {
                BoundingBoxXYZ bb = OOBB.GetOOBB(this.Solid);
                XYZ sizes = GetSizes(bb);

                _volume = sizes.X * sizes.Y * sizes.Z;

                _volume = UnitUtils.Convert(_volume, DisplayUnitType.DUT_CUBIC_FEET, DisplayUnitType.DUT_CUBIC_METERS); ;
                _volume = Math.Round(_volume, 15);

                return _volume;
            }
        }

        private XYZ GetSizes(BoundingBoxXYZ bb)
        {
            double x = bb.Max.X - bb.Min.X;
            double y = bb.Max.Y - bb.Min.Y;
            double z = bb.Max.Z - bb.Min.Z;
            double e;
            if (z > y) { e = z; z = y; y = e; }
            if (y > x) { e = y; y = x; x = e; }
            if (z > y) { e = z; z = y; y = e; }
            return new XYZ(x, y, z);
        }
    }

    public class ASPlate : SteelElement
    {
        private Plate _plate;

        private string _gost;
        private string _metal;
        private string _name;
        private double _mass;
        private string _uses;

        private double _volume;

        public override string GetGost()
        {
            GostNames.TryGetValue("t", out _gost);
            return _gost;
        }

        public override string GetMetal()
        {
            _metal = this.Plate.MaterialDescription;
            _metal = Regex.Replace(_metal.Split()[0], @"[^0-9a-zA-Z0-9А-Яа-я-\ ]+", " ") + "\nГОСТ27772";
            return _metal;
        }

        public override string GetName()
        {
            if (this.Plate.Thickness >= 25)
            {
                _name = "t" + this.Plate.Thickness + " z25";
            }
            else
            {
                _name = "t" + this.Plate.Thickness;
            }
            return _name;
        }

        public override double GetMass()
        {
            _mass = this.Volume * density;
            return _mass;
        }

        public override string GetUses()
        {
            _uses = this.Plate.MaterialDescription;
            _uses = _uses.Remove(0, _uses.IndexOf(' ') + 1);
            return _uses;
        }

        public Plate Plate
        {
            get { return _plate; }
            set { _plate = value; }
        }

        private double Volume
        {
            get
            {
                _volume = this.Plate.Width * this.Plate.Length * this.Plate.Thickness;
                _volume = UnitUtils.Convert(_volume, DisplayUnitType.DUT_CUBIC_MILLIMETERS, DisplayUnitType.DUT_CUBIC_METERS);
                _volume = Math.Round(_volume, 15);

                return _volume;
            }
        }

    }
}
