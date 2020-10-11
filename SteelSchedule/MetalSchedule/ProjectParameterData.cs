#region Namespaces
using Autodesk.Revit.DB;
#endregion

namespace SteelSchedule.MetalSchedule
{
    class ProjectParameterData
    {
        public Definition Definition = null;
        public ElementBinding Binding = null;
        public string Name = null;                // Needed because accsessing the Definition later may produce an error.
        public bool IsSharedStatusKnown = false;  // Will probably always be true when the data is gathered
        public bool IsShared = false;
        public string GUID = null;
    }
}
