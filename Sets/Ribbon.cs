using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Sets
{
    public partial class Ribbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void btnFindDuplicates_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Window window;
            Excel.Application application;
            Excel.Worksheet activeWorksheet;
            Excel.Range activeRange;
            _GetActiveSelections(e, out application, out window, out activeWorksheet, out activeRange);

            Excel.Range data = null;
            if(_ShowRangeSelector("Select data set", application, out data))
            {
                List<string> list = _BuildList(data);
                var duplicates = list.GroupBy(x => x).Where(g => g.Count() > 1).Select(g => new { Value = g.Key, Count = g.Count() }).ToList();

                int R = activeRange.Row;
                int C = activeRange.Column;
                int numDupes = duplicates.Count();
                MessageBox.Show("Found " + numDupes);
                Excel.Range current = null;
                for (int row = R, i = 0; i < numDupes; row++, i++)
                {
                    // write the value
                    current = activeRange.Worksheet.Cells[row, C];
                    current.Value2 = duplicates[i].Value;

                    // write the count next to the value
                    current = activeRange.Worksheet.Cells[row, C + 1];
                    current.Value2 = duplicates[i].Count;
                }
            }
        }

        private void btnUnion_Click(object sender, RibbonControlEventArgs e)
        {
            _Calculate(e, _CalculateUnion);
        }

        private void btnNotSubset_Click(object sender, RibbonControlEventArgs e)
        {
            _Calculate(e, _CalculateNotSubset);
        }

        private void btnNotIntersect_Click(object sender, RibbonControlEventArgs e)
        {
            _Calculate(e, _CalculateNotIntersect);
        }

        private void btnIntersection_Click(object sender, RibbonControlEventArgs e)
        {
            _Calculate(e, _CalculateIntersect);
        }

        private void _Calculate(RibbonControlEventArgs e, Func<List<string>, List<string>, List<string>> calculator)
        {
            Excel.Window window;
            Excel.Application application;
            Excel.Worksheet activeWorksheet;
            Excel.Range activeRange;
            _GetActiveSelections(e, out application, out window, out activeWorksheet, out activeRange);

            Excel.Range firstRange, secondRange;
            if (_ShowRangeSelector("Select first set", application, out firstRange) && _ShowRangeSelector("Select second set", application, out secondRange))
            {
                var list = calculator(_BuildList(firstRange), _BuildList(secondRange));

                MessageBox.Show("Found " + list.Count);
                _FillRange(activeRange, list);
            }
        }

        private void _GetActiveSelections(RibbonControlEventArgs e, out Excel.Application application, out Excel.Window window, out Excel.Worksheet activeWorksheet, out Excel.Range activeRange)
        {
            window = e.Control.Context;
            application = (Excel.Application)window.Application;
            activeWorksheet = (Excel.Worksheet)window.Application.ActiveSheet;
            activeRange = window.Application.ActiveCell;
        }

        private bool _ShowRangeSelector(string prompt, Excel.Application application, out Excel.Range range)
        {
            range = null;
            var set = application.InputBox(prompt, "Range Selector", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            if(set is Excel.Range)
            {
                range = (Excel.Range)set;
                return true;
            }
            return false;
        }

        private List<string> _BuildList(Excel.Range range)
        {
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            var list = new List<string>(rows * cols);
            for(int row = 1; row <= rows; row++)
            {
                for(int col = 1; col <= cols; col++)
                {
                    var temp = ((Excel.Range)range.Cells[row, col]);
                    var sz = (temp.Value2 == null) ? null : temp.Value2.ToString();
                    list.Add(sz);
                }
            }
            return list;
        }

        private void _FillRange(Excel.Range start, List<string> data)
        {
            int R = start.Row;
            int C = start.Column;
            Excel.Range current = null;
            for(int row = R, i = 0; i < data.Count; row++, i++)
            {
                current = start.Worksheet.Cells[row, C];
                current.Value2 = data[i];
            }
        }

        private List<string> _CalculateIntersect(List<string> first, List<string> second)
        {
            return new List<string>(first.Intersect(second));
        }

        private List<string> _CalculateNotIntersect(List<string> first, List<string> second)
        {
            return new List<string>(first.Except(second).Union(second.Except(first)));
        }

        private List<string> _CalculateNotSubset(List<string> first, List<string> second)
        {
            return new List<string>(first.Except(second));
        }

        private List<string> _CalculateUnion(List<string> first, List<string> second)
        {
            return new List<string>(first.Union(second).Distinct());
        }

        private void btnFileMerge_Click(object sender, RibbonControlEventArgs e)
        {
            List<string> files = null;
            using (var dialog = new OpenFileDialog())
            {
                dialog.Multiselect = true;
                dialog.Filter = "Excel (*.xlsx)|*.xlsx";
                var result = dialog.ShowDialog();
                if(result == DialogResult.OK)
                {
                    files = new List<string>(dialog.FileNames);
                }
            }

            if (files == null) return;

            Excel.Window window = (Excel.Window)e.Control.Context;
            Excel.Application application = window.Application;
            Excel.Workbook activeBook = application.ActiveWorkbook;
            application.ScreenUpdating = false;
            for (int i = 0; i < files.Count; i++)
            {
                string path = files[i];
                string filename = Path.GetFileNameWithoutExtension(path);
                Excel.Workbook book =  application.Workbooks.Open(path);
                Excel.Worksheet sheet = book.Sheets[1];
                sheet.Name = filename;
                sheet.Copy(Type.Missing, activeBook.Sheets[1]);
                book.Close(false);
            }
            application.ScreenUpdating = true;
        }
    }
}
