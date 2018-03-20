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
            Excel.Window window = e.Control.Context;
            Excel.Application application = (Excel.Application)window.Application;
            Excel.Worksheet activeWorksheet = (Excel.Worksheet)window.Application.ActiveSheet;
            Excel.Range activeRange = window.Application.ActiveCell;

            Excel.Range firstRange, secondRange;
            if (_GetRange("Select first set", application, out firstRange) && _GetRange("Select second set", application, out secondRange))
            {
                var list = calculator(_BuildList(firstRange), _BuildList(secondRange));

                MessageBox.Show("Found " + list.Count);
                _FillRange(activeRange, list);
            }
        }

        private bool _GetRange(string prompt, Excel.Application application, out Excel.Range range)
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
                    var sz = (string)temp.Value2;
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
