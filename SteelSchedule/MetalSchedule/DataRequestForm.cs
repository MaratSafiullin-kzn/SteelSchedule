using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

using RVTExternalCommandData = Autodesk.Revit.UI.ExternalCommandData;
using RVTDocument = Autodesk.Revit.DB.Document;
using RVTransaction = Autodesk.Revit.DB.Transaction;

using SteelSchedule.MetalSchedule;

namespace SteelSchedule
{
    public partial class DataRequestForm : Form
    {
        private List<SteelElement> steelelements;
        private RVTDocument _doc;
        private RVTExternalCommandData _commandData;

        public DataRequestForm(List<SteelElement> sEl, RVTExternalCommandData commandData)
        {
            InitializeComponent();
            unitComboBox.SelectedIndex = 0;
            steelelements = sEl;
            _commandData = commandData;
            _doc = _commandData.Application.ActiveUIDocument.Document;
            this.steelElemCount.Text = sEl.Count.ToString();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            try
            {
                using (RVTransaction t = new RVTransaction(_doc))
                {
                    t.Start("CreateSchedule");

                    double coefficient = double.Parse(this.coefficient.Text.Replace(".", ","));
                    Schedule sch = new Schedule();

                    sch.Create(this.unitComboBox.Text, coefficient, steelelements, _doc);

                    t.Commit();
                    this.Close();

                    _commandData.Application.ActiveUIDocument.ActiveView = sch.rvtSchedule;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message +": "+ ex.Source);
            }
        }
    }
}
