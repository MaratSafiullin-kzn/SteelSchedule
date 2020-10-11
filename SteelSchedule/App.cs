#region Namespaces
using System;
using System.IO;
using System.Collections.Generic;
using System.Collections;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Text;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

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

namespace SteelSchedule
{
    class App : IExternalApplication
    {
        static String addinAssmeblyPath = typeof(App).Assembly.Location;

        void createRibbonButton(UIControlledApplication application)
        {
            RibbonPanel panel;
            List<RibbonPanel> panelList = application.GetRibbonPanels();

            if (panelList.Where(x => x.Name == "АО Казанский ГИПРОНИИАВИАПРОМ").Any())
            {
                panel = panelList.Where(x => x.Name == "АО Казанский ГИПРОНИИАВИАПРОМ").First();
            }
            else
            {
                panel = application.CreateRibbonPanel("АО Казанский ГИПРОНИИАВИАПРОМ");
            }

            PushButtonData pbd_Options = new PushButtonData("MetalSchedule", "Спецификация \n металлопроката", addinAssmeblyPath, "SteelSchedule.MetalSchedule.MetalSchedulePushButton");

            pbd_Options.LargeImage = convertFromBitmap(SteelSchedule.Properties.Resources.imageMetalSchegule);
            pbd_Options.LongDescription = "Спецификация металлопроката";

            Autodesk.Revit.UI.PushButton pb_Options = panel.AddItem(pbd_Options) as Autodesk.Revit.UI.PushButton;
        }

        BitmapSource convertFromBitmap(System.Drawing.Bitmap bitmap)
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                bitmap.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }

        public Result OnStartup(UIControlledApplication a)
        {
            createRibbonButton(a);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
