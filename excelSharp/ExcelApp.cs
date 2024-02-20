using System;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;


namespace excelSharp
{
    internal class ExcelApp
    {
        Excel.Application oXL;
        Excel._Workbook oWB;
        public ExcelApp()
        {
            oXL = new Excel.Application();

        }
        ~ExcelApp()
        {
            if (oXL != null)
            {
                if (oWB != null)
                {
                    try
                    {
                        oWB.Close(0);
                    }
                    catch (Exception) { }
                }

                oXL.Quit();
            }
        }
        public TimeTable readTimeTable(string file)
        {
            Excel._Worksheet oSheet;
            Range r;

            int[] paresCount;
            string[] groups = { };
            int[] subgroups = { };

            try
            {
                oWB = (Excel._Workbook)(oXL.Workbooks.Open(file, Type.Missing, true));
                oSheet = (Excel._Worksheet)(oWB.ActiveSheet);

                //считаем число пар по дням
                paresCount = getParesCount(oSheet);
                getGroupList(oSheet, ref groups, ref subgroups);


                List<GroupTimeTable> timeTable = readTimeTable(oSheet, subgroups, paresCount);


                Dictionary<string, Group> groupList = new Dictionary<string, Group>();
                int i = 0;
                int column = 0;
                foreach (var groupName in groups)
                {
                    Group group = new Group(groupName, subgroups[i]);
                    groupList.Add(groupName, group);
                    for(int j = 0; j < subgroups[i]; j++)
                    {
                        var tmp = timeTable[column++];
                        group.addTimetable(tmp);
                    }
                    i++;
                }
                oWB.Close();
                return new TimeTable(groupList, paresCount);

            }
            catch(Exception ex) {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, ex.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, ex.Source);

                MessageBox.Show(errorMessage, "Error");
            }
            return null;
        }
        private List<GroupTimeTable> readTimeTable(_Worksheet sheet, int[] subgroups, int[] paresCount)
        {
            int columnCount = 0;
            foreach (int group in subgroups)
            {
                columnCount += group;
            }
            int RowsCount = 0;
            foreach (int row in paresCount)
            {
                RowsCount += row * 2;
            }

            Range r;

            List<GroupTimeTable> timtable = new List<GroupTimeTable>();
            GroupTimeTable groupTimtable;

            bool anyWeek;

            for (int i = 0, column = 3, row = 12; i < columnCount; i++, column++)
            {
                groupTimtable = new GroupTimeTable();
                for (int j = 0; j < RowsCount; j++)
                {
                    r = sheet.Cells[row, column];
                    anyWeek = r.MergeArea.Rows.Count > 2;
                    if (r.MergeArea.Columns.Count > 1)
                    {
                        normolizeTable(sheet, r, columnCount, column, row, anyWeek);
                        r = sheet.Cells[row, column];
                    }
            
                    if(anyWeek)
                    {
                        groupTimtable.Add(r.Value);
                        
                    }
                    else
                    {
                        groupTimtable.Add(r.Value, r.Offset[1, 0].Value);
                    }

                    row += 4;
                    j++;
                }
                timtable.Add(groupTimtable);
                row = 12;
            }
            return timtable;
        }
        private void normolizeTable(_Worksheet sheet, Range r, int maxColumns, int column, int row, bool merge)
        {
            int columns = r.MergeArea.Columns.Count;
            string value = r.Value;
            r.UnMerge();
            Range newRange;
            for(int i = 0; i < maxColumns && i < columns; i++)
            {
                newRange = sheet.Cells[row, column + i];
                if(merge)
                {
                    newRange = sheet.get_Range(newRange.Address, newRange.Offset[2, 0].Columns.Address);
                    newRange.Merge();
                }
                newRange.Value = value;

            }
        }
        private void getGroupTimeTable(_Worksheet sheet, int group)
        {

        }
        private void getGroupList(_Worksheet sheet, ref string[] groups, ref int[] subgroups)
        {
            List<string> groupsName = new List<string>();
            List<int> subGroupsList = new List<int>();
            int count = 0;
            Range r;
            int shift = 3;
            r = sheet.Cells[10, shift];

            while (r.Value != null && r.Value != "")
            {
                count = r.MergeArea.Columns.Count;
                shift += count;
                subGroupsList.Add(count);
                groupsName.Add(r.Value);
                r = sheet.Cells[10, shift];
            }
            groups = groupsName.ToArray();
            subgroups = subGroupsList.ToArray();

        }
        private int[] getParesCount(_Worksheet sheet)
        {
            Range r;
            int[] paresCount = { 0, 0, 0, 0, 0, 0 };
            int shift = 12;

            int count;
            for (int i = 0; i < paresCount.Length; i++)
            {
                r = sheet.Cells[shift, 1];
                if (r.MergeCells)
                {
                    count = r.MergeArea.Count;
                    shift += count;
                    paresCount[i] = count / 4;
                }
            }

            return paresCount;
        }
        public List<string> readStudentsFromExcel(string file)
        {
            List<string> list = new List<string>();
            Excel._Worksheet oSheet;
            Range r;
            try
            {
                //oXL.Visible = true;
                oWB = (Excel._Workbook)(oXL.Workbooks.Open(file, Type.Missing, true));
                oSheet = (Excel._Worksheet)(oWB.ActiveSheet);
                
                list.Add(oWB.Name.Substring(0, oWB.Name.Length - 5));
                
                string student;
                int i = 1;
                r = oSheet.Cells[i, 1];
                student = r.Value;
                while(student != null && student.Length > 1)
                {
                    list.Add(student);
                    r = oSheet.Cells[++i, 1];
                    student = r.Value;
                }


                oWB.Close();
            }
            catch (Exception ex)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, ex.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, ex.Source);

                MessageBox.Show(errorMessage, "Error");
            }
            return list;
        }
        public void createTable(List<string> students)
        {
            if(oXL == null)
            {
                oXL = new Excel.Application();
            }

            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                if (oWB != null)
                {
                    oWB.Close(0);

                }
                oXL.Visible = true;



                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                createTemplateTable(students.Count, oSheet);

                addStudentsToTable(students, oSheet);

                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception ex)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, ex.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, ex.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }
        private void addStudentsToTable(List<string> students, _Worksheet oSheet)
        {
            int i = 1;
            foreach (var student in students)
            {
                oSheet.Cells[5 + i++, 2].value = student;
            } 
        }
        private void createTemplateTable(int studCount, _Worksheet oSheet)
        {
            Range range = oSheet.get_Range("A1", "B1");
            range.Merge();
            range.Cells[1, 1] = "Дни недели - Числитель";

            range = oSheet.get_Range("A2", "B2");
            range.Merge();
            range.Cells[1, 1] = "Дата Занятий";

            range = oSheet.get_Range("AM1", "AN2");
            range.Merge();
            range.Value = "Всего пропущенно занятий";
            range.EntireColumn.ColumnWidth = 6.63;
            setFontToCenter(range);
            range.WrapText = true;

            range = oSheet.get_Range("AM3", "AM4");
            range.Merge();
            range.Value = "По уважительным";
            range.WrapText = true;
            range.Font.Size = 8;

            range = oSheet.get_Range("AN3", "AN4");
            range.Merge();
            range.Value = "По не уважительным";
            range.WrapText = true;
            range.Font.Size = 8;

            int lastCell = studCount + 5 + 2;

            range = oSheet.get_Range("4:4");
            range.EntireRow.RowHeight = 22.5;
            range = oSheet.get_Range("6:" + (lastCell - 1));
            range.EntireRow.RowHeight = 23;

            range = oSheet.get_Range(lastCell + ":" + lastCell);
            range.EntireRow.RowHeight = 55;

            range = oSheet.get_Range("B" + lastCell);
            range.Value = "подпись преподавателей";
            setFontToCenter(range);
            range.VerticalAlignment = XlHAlign.xlHAlignCenter;

            range = oSheet.get_Range("A3", "B4");
            range.Merge();
            range.Cells[1, 1] = "Вид Занятий";
            setFontToCenter(range);


            range = oSheet.get_Range("B:B");
            range.EntireColumn.ColumnWidth = 43.5;
            range = oSheet.get_Range("5:5");
            range.EntireRow.RowHeight = 83.25;
            oSheet.get_Range("B5").Borders[XlBordersIndex.xlDiagonalDown].Weight = XlBorderWeight.xlThin;

            range = oSheet.get_Range("A:A");
            range.EntireColumn.ColumnWidth = 3;



            range = oSheet.get_Range("A1", "AN" + lastCell);
            range.Borders.Weight = XlBorderWeight.xlThin;


            range = oSheet.get_Range("A1", "AL4");
            setFontToCenter(range);
            range = oSheet.get_Range("C1", "AL2");
            range.EntireColumn.ColumnWidth = 4;
            range.Font.Size = 12;

            createLabel(oSheet, "Дисциплина", 170.0f, 90.0f);
            createLabel(oSheet, "Ф.И.О. студента", 40.0f, 120.0f);


            range = oSheet.get_Range("C1", "AL1");
            string[] days =
            {
                    "Понедельник",
                    "Вторник",
                    "Среда",
                    "Четверг",
                    "Пятница",
                    "Суббота"
                };

            setRangeShiftText(oSheet, range, 6, days);


            setMergeText(oSheet.get_Range("C2", "H2"), "05.02.2024");

            range = oSheet.get_Range("C2", "AL2");
            range.NumberFormat = "DD/MM/YYYY";

            int j = 0;
            Range last = range.Columns;
            string text;
            foreach (Range item in range.Columns)
            {
                if (j == 0)
                {
                    last = item;

                }
                else
                {
                    if (j % 6 == 0)
                    {
                        text = "=" + last.Address + " + 1";
                        last = item;
                        setMergeText(oSheet.get_Range(item.Address, item.Offset[0, 5].Columns.Address), text);
                    }

                }
                j++;
            }


            range = oSheet.get_Range("C3", "AL5");
            range.Orientation = 90.0;
            setFontToCenter(range);

            range = oSheet.get_Range("C3", "AL4");
            foreach (Range item in range.Columns)
            {
                oSheet.get_Range(item.Address).Merge();
            }






            for (int i = 1, length = studCount; i <= length; i++)
            {
                oSheet.Cells[5 + i, 1].value = i;
            }

            oSheet.get_Range("A" + lastCell, "AN" + lastCell).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThick;
        }

        private void createLabel(_Worksheet sheet, string text, float x, float y)
        {
            Shape shape = sheet.Shapes.AddLabel(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast,
                    x, y, 5.0f, 10.0f);
            shape.TextFrame2.TextRange.Font.Size = 11;
            shape.TextFrame2.TextRange.Text = text;
            shape.TextFrame2.Orientation = Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal;
        }
        private void setFontToCenter(Range range)
        {
            range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }
        private void setMergeText(Range range, string text)
        {
            range.Merge();
            range.Cells[1, 1] = text;
        }
        private void setRangeShiftText(_Worksheet oSheet, Range range, int shift, string[] values)
        {
            int j = 0;
            foreach (Range item in range.Columns)
            {
                if (j % shift == 0)
                {
                    setMergeText(oSheet.get_Range(item.Address, item.Offset[0, 5].Columns.Address), values[j / shift]);
                }
                j++;
            }
        }
        private void DisplayQuarterlySales(Excel._Worksheet oWS)
        {
            //Excel._Workbook oWB;
            //Excel.Series oSeries;
            //Excel.Range oResizeRange;
            //Excel._Chart oChart;
            //String sMsg;
            //int iNumQtrs;

            ////Determine how many quarters to display data for.
            //for (iNumQtrs = 4; iNumQtrs >= 2; iNumQtrs--)
            //{
            //    sMsg = "Enter sales data for ";
            //    sMsg = String.Concat(sMsg, iNumQtrs);
            //    sMsg = String.Concat(sMsg, " quarter(s)?");

            //    DialogResult iRet = MessageBox.Show(sMsg, "Quarterly Sales?",
            //    MessageBoxButtons.YesNo);
            //    if (iRet == DialogResult.Yes)
            //        break;
            //}

            //sMsg = "Displaying data for ";
            //sMsg = String.Concat(sMsg, iNumQtrs);
            //sMsg = String.Concat(sMsg, " quarter(s).");

            //MessageBox.Show(sMsg, "Quarterly Sales");

            ////Starting at E1, fill headers for the number of columns selected.
            //oResizeRange = oWS.get_Range("E1", "E1").get_Resize(Missing.Value, iNumQtrs);
            //oResizeRange.Formula = "=\"Q\" & COLUMN()-4 & CHAR(10) & \"Sales\"";

            ////Change the Orientation and WrapText properties for the headers.
            //oResizeRange.Orientation = 38;
            //oResizeRange.WrapText = true;

            ////Fill the interior color of the headers.
            //oResizeRange.Interior.ColorIndex = 36;

            ////Fill the columns with a formula and apply a number format.
            //oResizeRange = oWS.get_Range("E2", "E6").get_Resize(Missing.Value, iNumQtrs);
            //oResizeRange.Formula = "=RAND()*100";
            //oResizeRange.NumberFormat = "$0.00";

            ////Apply borders to the Sales data and headers.
            //oResizeRange = oWS.get_Range("E1", "E6").get_Resize(Missing.Value, iNumQtrs);
            //oResizeRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

            ////Add a Totals formula for the sales data and apply a border.
            //oResizeRange = oWS.get_Range("E8", "E8").get_Resize(Missing.Value, iNumQtrs);
            //oResizeRange.Formula = "=SUM(E2:E6)";
            //oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle
            //= Excel.XlLineStyle.xlDouble;
            //oResizeRange.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight
            //= Excel.XlBorderWeight.xlThick;

            ////Add a Chart for the selected data.
            //oWB = (Excel._Workbook)oWS.Parent;
            //oChart = (Excel._Chart)oWB.Charts.Add(Missing.Value, Missing.Value,
            //Missing.Value, Missing.Value);

            ////Use the ChartWizard to create a new chart from the selected data.
            //oResizeRange = oWS.get_Range("E2:E6", Missing.Value).get_Resize(
            //Missing.Value, iNumQtrs);
            //oChart.ChartWizard(oResizeRange, Excel.XlChartType.xl3DColumn, Missing.Value,
            //Excel.XlRowCol.xlColumns, Missing.Value, Missing.Value, Missing.Value,
            //Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //oSeries = (Excel.Series)oChart.SeriesCollection(1);
            //oSeries.XValues = oWS.get_Range("A2", "A6");
            //for (int iRet = 1; iRet <= iNumQtrs; iRet++)
            //{
            //    oSeries = (Excel.Series)oChart.SeriesCollection(iRet);
            //    String seriesName;
            //    seriesName = "=\"Q";
            //    seriesName = String.Concat(seriesName, iRet);
            //    seriesName = String.Concat(seriesName, "\"");
            //    oSeries.Name = seriesName;
            //}

            //oChart.Location(Excel.XlChartLocation.xlLocationAsObject, oWS.Name);

            ////Move the chart so as not to cover your data.
            //oResizeRange = (Excel.Range)oWS.Rows.get_Item(10, Missing.Value);
            //oWS.Shapes.Item("Chart 1").Top = (float)(double)oResizeRange.Top;
            //oResizeRange = (Excel.Range)oWS.Columns.get_Item(2, Missing.Value);
            //oWS.Shapes.Item("Chart 1").Left = (float)(double)oResizeRange.Left;
        }
    }
}
