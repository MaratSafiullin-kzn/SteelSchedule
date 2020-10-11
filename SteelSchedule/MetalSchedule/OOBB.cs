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
    public class OOBB
    {
        static XYZ Min(XYZ a, XYZ b)
        {
            double x, y, z;
            if (a.X < b.X) { x = a.X; } else { x = b.X; }
            if (a.Y < b.Y) { y = a.Y; } else { y = b.Y; }
            if (a.Z < b.Z) { z = a.Z; } else { z = b.Z; }
            return new XYZ(x, y, z);
        }

        static XYZ Max(XYZ a, XYZ b)
        {
            double x, y, z;
            if (a.X > b.X) { x = a.X; } else { x = b.X; }
            if (a.Y > b.Y) { y = a.Y; } else { y = b.Y; }
            if (a.Z > b.Z) { z = a.Z; } else { z = b.Z; }
            return new XYZ(x, y, z);
        }

        public static BoundingBoxXYZ GetOOBB(Solid solid)
        {
            IList<XYZ> vertices = new List<XYZ>();
            IList<PlanarFace> faces = new List<PlanarFace>();

            BoundingBoxXYZ AABB = new BoundingBoxXYZ();

            foreach (Face face in solid.Faces)
            {
                // filter only planar faces
                if (face is PlanarFace) { faces.Add(face as PlanarFace); }

                Mesh mesh = face.Triangulate();

                for (int i = 0; i < mesh.NumTriangles; i++)
                {
                    MeshTriangle triangle = mesh.get_Triangle(i);
                    XYZ p = triangle.get_Vertex(0); vertices.Add(p);
                    XYZ q = triangle.get_Vertex(1); vertices.Add(q);
                    XYZ r = triangle.get_Vertex(2); vertices.Add(r);

                    AABB.Min = Min(AABB.Min, p); AABB.Max = Max(AABB.Max, p);
                    AABB.Min = Min(AABB.Min, q); AABB.Max = Max(AABB.Max, q);
                    AABB.Min = Min(AABB.Min, r); AABB.Max = Max(AABB.Max, r);
                }
            }


            faces.OrderByDescending(f => f.Area);


            BoundingBoxXYZ minBB = new BoundingBoxXYZ();

            int count = 0;
            XYZ sizes = AABB.Max - AABB.Min;
            double volume = sizes.X * sizes.Y * sizes.Z;
            bool lessAABB = false;

            foreach (Face face in faces)
            {
                // Some optimization
                if (count > 24) break;
                count++;

                Transform tf = face.ComputeDerivatives(new UV()).Inverse;

                BoundingBoxXYZ bb = new BoundingBoxXYZ();
                bb.Min = new XYZ(1e10, 1e10, 1e10);
                bb.Max = new XYZ(-1e10, -1e10, -1e10);

                foreach (XYZ v in vertices)
                {
                    XYZ tv = tf.OfPoint(v);
                    bb.Min = Min(bb.Min, tv);
                    bb.Max = Max(bb.Max, tv);
                }

                XYZ sizes_ = bb.Max - bb.Min;
                double volume_ = sizes_.X * sizes_.Y * sizes_.Z;

                if (volume_ < volume)
                {
                    volume = volume_;
                    minBB = bb;
                    lessAABB = true;
                }
            }

            if (lessAABB) { return minBB; } else { return AABB; }
        }
    }
}
