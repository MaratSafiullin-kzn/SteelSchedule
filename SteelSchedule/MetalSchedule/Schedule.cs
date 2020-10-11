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
    class Schedule
    {
        private RVTDocument _doc;
        private ViewSchedule _schedule;

        private RVTDocument SetDoc
{
            set { _doc = value; }
        }

        public ViewSchedule rvtSchedule
        {
            get { return _schedule; }
            set { _schedule = value; }
        }

        public bool Create(string unit, double coefficient, List<SteelElement> _elementList, RVTDocument _doc)
        {
            try
            {
                this.SetDoc = _doc;

                double unitInd = 1;

                if (unit == "кг"){
                    unitInd = 1 * coefficient; }
                else if (unit == "т") {
                    unitInd = 0.001 * coefficient; }

                #region Сортировка
                var steelElementsMassSumByGNM = _elementList
                    .GroupBy(x => new
                            {
                                x.Gost, x.Name, x.Metal
                            })
                    .Select(x => new
                            {
                                _gost = x.Key.Gost,
                                _name = x.Key.Name,
                                _metal = x.Key.Metal,
                                _uses = x.Select(y => y.Uses),
                                _mass = Math.Round(x.Sum(y => y.Mass) * unitInd, 2)
                            })
                    .OrderBy(o => o._gost).ThenBy(o => o._metal).ThenBy(o => o._name)
                    .ToList();

                var steelElementsMassSumByGNMU = _elementList
                        .GroupBy(x => new
                        {
                            x.Gost,
                            x.Name,
                            x.Metal,
                            x.Uses
                        })
                        .Select(x => new
                        {
                            _gost = x.Key.Gost,
                            _name = x.Key.Name,
                            _metal = x.Key.Metal,
                            _uses = x.Key.Uses,
                            _mass = Math.Round(x.Sum(y => y.Mass) * unitInd, 2)
                        })
                        .OrderBy(o => o._gost).ThenBy(o => o._metal).ThenBy(o => o._name)
                        .ToList();
                #endregion

                //Получаем список всех назначений конструкций и назначаем индекс колонны. Индекс 4 стартовый для данной спецификации.
                Dictionary<string, int> uses = new Dictionary<string, int>();
                int UsesIndex = 4;
                foreach (var elem in steelElementsMassSumByGNM)
                {
                    foreach (var e in elem._uses)
                    {
                        if (uses.ContainsKey(e.ToString())) continue;
                        uses.Add(e.ToString(), UsesIndex);
                        UsesIndex++;
                    }
                }

                ViewSchedule schedule = CreateRevitSchedule(uses.Count);

                TableData colTableData = schedule.GetTableData();
                TableSectionData tsd = colTableData.GetSectionData(SectionType.Header);

                int colCount = 5 + uses.Count; // 5 это: первые четыре постоянные столбцы + пятый "всего".

                CreateColumnsAndRows(colCount, steelElementsMassSumByGNM.Count, tsd);
                CreateHeading(tsd, 4, uses, unit);

                //Запись значений в спецификацию Revit
                try
                {
                    int rowInd = 2;
                    foreach (var elem in steelElementsMassSumByGNM)
                    {
                        tsd.SetCellText(rowInd, 0, elem._gost);
                        tsd.SetCellText(rowInd, 1, elem._metal);
                        tsd.SetCellText(rowInd, 2, elem._name);

                        var qqq = steelElementsMassSumByGNMU.Where(_ => _._name == elem._name & _._metal == elem._metal).ToList();

                        elem._uses.Join(qqq, u => u, q => q._uses, (u, q) => new { u, q }).ToList().ForEach(_ =>
                        {
                            int usesColumn;
                            uses.TryGetValue(_.q._uses, out usesColumn);
                            tsd.SetCellText(rowInd, usesColumn, _.q._mass.ToString());
                        });
                        int colLastIndex = colCount - 1;
                        tsd.SetCellText(rowInd, colLastIndex, elem._mass.ToString());
                        rowInd++;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Запись значений в спецификацию Revit: " + ex.Message + " Source: " + ex.Source + " StackTrace:" + ex.StackTrace);
                }

                MergeAndSumm(tsd, uses);

                Text(tsd, coefficient);

                Numerator(tsd);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Основной: " + ex.Message + " Source: " + ex.Source + " StackTrace:" + ex.StackTrace);
                return false;
            }
        }

        private ViewSchedule CreateRevitSchedule(int usesCount)
        {
            ViewSchedule schedule = ViewSchedule.CreateSchedule(_doc, new ElementId(BuiltInCategory.OST_Windows), ElementId.InvalidElementId);

            string name = "Спецификация металлопроката  " + Regex.Replace(DateTime.Now.ToString(), ":", "-");
            schedule.Name = name;

            ScheduleDefinition definition = schedule.Definition;

            SchedulableField schedulableField =
                    definition.GetSchedulableFields().FirstOrDefault<SchedulableField>(sf => sf.ParameterId == new ElementId(BuiltInParameter.WINDOW_HEIGHT));

            ScheduleField field = definition.AddField(schedulableField);

            ScheduleFilter filter = new ScheduleFilter(field.FieldId, ScheduleFilterType.Equal, (double)12);
            definition.AddFilter(filter);

            TableData colTableData = schedule.GetTableData();
            TableSectionData tsdSS = colTableData.GetSectionData(SectionType.Body);

            double scw = 0.410104997 + usesCount * 0.049212598;

            tsdSS.SetColumnWidth(0, scw);

            this.rvtSchedule = schedule;
            return schedule;
        }

        private bool CreateColumnsAndRows(int columnCount, int rowsCount, TableSectionData tsd)
        {
            tsd.RemoveRow(0);
            tsd.RemoveColumn(0);

            for (int i = 0; i < columnCount; i++)
            {
                tsd.InsertColumn(i);
                tsd.SetColumnWidth(i, 0.049212598);


                if (i < 3)
                {
                    tsd.SetColumnWidth(i, 0.0984252);
                } //30 mm

                if (i == 3)
                {
                    tsd.SetColumnWidth(i, 0.0328084);
                } //10 mm

                if (i == columnCount - 1)
                {
                    tsd.SetColumnWidth(i, 0.082020997);
                } //25 mm
            }

            for (int i = 0; i < rowsCount + 1; i++)
            {
                tsd.InsertRow(i);
                tsd.SetRowHeight(i, 0.02);
            }

            return true;
        }

        private bool CreateHeading(TableSectionData tsd, int startUsesIndex, Dictionary<string, int> uses, string unit)
        {
            TableMergedCell tmc = new TableMergedCell();
            tsd.InsertRow(0);

            tmc.Top = 0;
            tmc.Bottom = 1;

            for (int i = 0; i < 4; i++)
            {
                tmc.Left = i;
                tmc.Right = i;

                tsd.MergeCells(tmc);
            }

            tmc.Left = tsd.LastColumnNumber;
            tmc.Right = tsd.LastColumnNumber;

            tsd.MergeCells(tmc);

            tmc.Top = 0;
            tmc.Bottom = 0;

            tmc.Left = startUsesIndex;
            tmc.Right = startUsesIndex + uses.Count - 1;

            tsd.MergeCells(tmc);

            tsd.SetCellText(0, 0, "Наименование профиля ГОСТ, ТУ");
            tsd.SetCellText(0, 1, "Марка металла");
            tsd.SetCellText(0, 2, "Номер или размеры профиля, мм");
            tsd.SetCellText(0, 3, "№, пп");
            tsd.SetCellText(0, 4, "Масса металла по элементам конструкции");
            tsd.SetCellText(0, tsd.LastColumnNumber, "Общая масса, " + unit);

            int colInd = 4;
            foreach (KeyValuePair<string, int> y in uses)
            {
                tsd.SetCellText(1, colInd, y.Key);
                colInd++;
            }

            return true;
        }

        private bool MergeAndSumm(TableSectionData tsd, Dictionary<string, int> uses)
        {
            int gostCol = 0;
            int metalCol = 1;
            double summ = 0;

            int s = tsd.LastRowNumber;

            #region Строка Всего
            for (int nRow = 2; nRow < s; nRow++)
            {
                if (tsd.GetCellText(nRow + 1, gostCol) != tsd.GetCellText(nRow, gostCol))
                {
                    InsertRowVsego(tsd, nRow);

                    nRow++;
                    s++;
                }

                if (s - 1 == nRow)
                {
                    InsertRowVsego(tsd, nRow + 1);
                    InsertRowItogo(tsd, nRow + 1);

                }
            }
            #endregion

            #region Сторока Итого
            for (int nRow = 2; nRow < s; nRow++)
            {
                if (tsd.GetCellText(nRow + 1, metalCol) != tsd.GetCellText(nRow, metalCol))
                {
                    if (tsd.GetCellText(nRow + 1, metalCol) == "")
                    {
                        InsertRowItogo(tsd, nRow);
                        s++;

                        nRow = nRow + 2;
                    }
                    else
                    {
                        InsertRowItogo(tsd, nRow);
                        s++;

                        nRow++;
                    }
                }
            }
            #endregion

            #region Объединение яцеек
            int startMergeGost = 2;
            int startMergeMetal = 2;

            s = s + 2;
            for (int nRow = 2; nRow < s; nRow++)
            {
                if (tsd.GetCellText(nRow + 1, gostCol) != tsd.GetCellText(nRow, gostCol))
                {
                    MergeScheduleCells(tsd, gostCol, startMergeGost, nRow);

                    startMergeGost = nRow + 1;
                }

                if (tsd.GetCellText(nRow + 1, metalCol) != tsd.GetCellText(nRow, metalCol))
                {
                    MergeScheduleCells(tsd, metalCol, startMergeMetal, nRow);

                    startMergeMetal = nRow + 1;
                }
            }
            #endregion

            #region Суммирование яцеек
            int startSummGost = 2;
            int startSummMetal = 2;

            for (int nRow = 1; nRow < s + 1; nRow++)
            {
                #region Всего
                if (tsd.GetCellText(nRow, gostCol) == "Всего")
                {
                    for (int r = 4; r < tsd.LastColumnNumber + 1; r++)
                    {
                        for (int i = startSummGost; i < nRow; i++)
                        {
                            if (tsd.GetCellText(i, r) != "")
                            {
                                if (tsd.GetCellText(i, metalCol) != "Итого")
                                {
                                    summ = summ + Convert.ToDouble(tsd.GetCellText(i, r));
                                }
                            }
                        }

                        if (summ != 0)
                        {
                            tsd.SetCellText(nRow, r, summ.ToString());
                        }

                        summ = 0;
                    }

                    startSummGost = nRow + 1;
                }
                #endregion

                #region Итого
                if (tsd.GetCellText(nRow, metalCol) == "Итого")
                {
                    for (int r = 4; r < tsd.LastColumnNumber + 1; r++)
                    {
                        for (int i = startSummMetal; i < nRow; i++)
                        {
                            if (tsd.GetCellText(i, r) != "")
                            {
                                summ = summ + Convert.ToDouble(tsd.GetCellText(i, r));
                            }
                        }

                        if (summ != 0)
                        {
                            tsd.SetCellText(nRow, r, summ.ToString());
                        }

                        summ = 0;
                    }

                    if (tsd.GetCellText(nRow + 1, gostCol) == "Всего")
                    {
                        startSummMetal = nRow + 2;
                    }
                    else
                    {
                        startSummMetal = nRow + 1;
                    }
                }
                #endregion
            }
            #endregion

            #region Всего масса металла
            InsertRowVsegoMasMet(tsd, tsd.LastRowNumber);

            double startSummMetalVsego = 0;

            for (int nCol = 4; nCol < tsd.LastColumnNumber + 1; nCol++)
            {
                for (int nRow = 2; nRow < s + 1; nRow++)
                {
                    if (tsd.GetCellText(nRow, nCol) != "")
                    {
                        if (tsd.GetCellText(nRow, metalCol) != "Итого")
                        {
                            if (tsd.GetCellText(nRow, gostCol) != "Всего")
                            {
                                startSummMetalVsego = startSummMetalVsego + Convert.ToDouble(tsd.GetCellText(nRow, nCol));
                            }
                        }
                    }
                }

                tsd.SetCellText(tsd.LastRowNumber, nCol, startSummMetalVsego.ToString());
                startSummMetalVsego = 0;
            }

            #endregion

            return true;
        }

        private bool InsertRowVsego(TableSectionData tsd, int nRow)
        {
            tsd.InsertRow(nRow + 1);
            tsd.SetRowHeight(nRow + 1, 0.02);

            tsd.SetCellText(nRow + 1, 0, "Всего");

            return true;
        }

        private bool InsertRowVsegoMasMet(TableSectionData tsd, int nRow)
        {
            tsd.InsertRow(nRow + 1);
            tsd.SetRowHeight(nRow + 1, 0.02);

            tsd.SetCellText(nRow + 1, 0, "Всего масса металла");

            return true;
        }

        private bool InsertRowItogo(TableSectionData tsd, int nRow)
        {
            tsd.InsertRow(nRow + 1);
            tsd.SetRowHeight(nRow + 1, 0.02);

            tsd.SetCellText(nRow + 1, 1, "Итого");
            tsd.SetCellText(nRow + 1, 0, tsd.GetCellText(nRow, 0));

            return true;
        }

        private bool MergeScheduleCells(TableSectionData tsd, int columnIndex, int top, int botton)
        {
            TableMergedCell tmc = new TableMergedCell();

            tmc.Left = columnIndex;
            tmc.Right = columnIndex;

            tmc.Top = top;
            tmc.Bottom = botton;

            tsd.MergeCells(tmc);

            return true;
        }

        private bool Text(TableSectionData tsd, double multiplier)
        {

            tsd.InsertRow(tsd.LastRowNumber + 1);

            TableMergedCell tmc = new TableMergedCell();
            tmc.Left = 0;
            tmc.Right = tsd.LastColumnNumber;

            tmc.Top = tsd.LastRowNumber;
            tmc.Bottom = tsd.LastRowNumber;

            tsd.MergeCells(tmc);

            tsd.SetCellText(tsd.LastRowNumber, 0, "На всю спецификацию учтен коэффикиент: " + multiplier.ToString());

            return true;
        }

        private bool Numerator(TableSectionData tsd)
        {
            int rowIndex = 1;
            for (int nRow = 2; nRow < tsd.LastRowNumber + 1; nRow++)
            {
                tsd.SetCellText(nRow, 3, rowIndex.ToString());
                rowIndex++;
            }
            return true;
        }
    }

}


