using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Syroot.Windows.IO;

namespace Excel_Export
{
    public class Export
    {
        public string FolderLocation { get; set; }
        public string Suffix { get; set; }

        public List<Result> results { get; set; }

        private int _ColumnIndex;
        public int ColumnIndex
        {
            get { return this._ColumnIndex; }
            set
            {
                if (value <= TotalCol && value > 0)
                    this._ColumnIndex = value;
            }
        }

        private int _TableHeader;
        public int TableHeader
        {
            get { return this._TableHeader; }
        set
            {
                if (value <= TotalRow && value > 0)
                    this._TableHeader = value;
            }
        }

        public HashSet<string> UniqueList { get; set; }
        public int TotalRow { get; set; }
        public int TotalCol { get; set; }

        private Microsoft.Office.Interop.Excel.Application xlApp;
        private Microsoft.Office.Interop.Excel.Worksheet ws;
        private object[,] array;


        public Export(Microsoft.Office.Interop.Excel.Application application)
        {
            this.xlApp = application;
            this.ws = xlApp.ActiveSheet as Worksheet;

            this.TotalCol = ws.UsedRange.Columns.Count;
            this.TotalRow = ws.UsedRange.Rows.Count;

            this.FolderLocation = new KnownFolder(KnownFolderType.Downloads).Path.ToString();
            this.Suffix = DateTime.Today.ToString("yyyy.MM.dd") + "_";
            this.ColumnIndex = 1;
            this.TableHeader = 2;

            ws.EnableAutoFilter = true;

            this.array = ws.UsedRange.Value;

            UniqueList = new HashSet<string>();
            UpdateUniqueList();

        }

        public void UpdateUniqueList()
        {
            if (array != null)
            {
                UniqueList.Clear();

                for (int i = TableHeader; i <= TotalRow; i++)
                {
                    if (i <= array.Length)
                    {
                        var cell = array[i, ColumnIndex];
                        if (cell != null)
                        {
                            string part = cell.ToString();
                            if (part != "")
                            {
                                UniqueList.Add(part);
                            }
                        }
                    } 
                }
            }
        }

        public void ToNewFiles()
        {

            results = new List<Result>();

            foreach (var item in UniqueList)
            {
                Range range = ws.UsedRange;
                range.AutoFilter(Field: ColumnIndex, Criteria1: item, VisibleDropDown: true);

                while (xlApp.CalculationState != XlCalculationState.xlDone)
                {
                    Task.Delay(25);
                }

                Range from = ws.UsedRange;

                Workbook newbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                Worksheet nws = newbook.Worksheets[1];
                Range dest = nws.Range["A1"];

                //Copy funziona solo se xlApp è la medesima!!!
                from.Copy(dest);

                Result r = new Result { PageName = item, RowCount = ws.AutoFilter.Range.Columns[ColumnIndex].SpecialCells(XlCellType.xlCellTypeVisible).Cells.Count - (TableHeader - 1)};
                results.Add(r);

                newbook.SaveAs(FolderLocation + "\\" + Suffix + item + ".xlsx");
                newbook.Close();
            }
        }

        public void ToNewSheets()
        {

            results = new List<Result>();

            foreach (var item in UniqueList)
            {
                Range range = ws.UsedRange;
                range.AutoFilter(Field: ColumnIndex, Criteria1: item, VisibleDropDown: true);
                Range from = ws.UsedRange;

                Worksheet newWorksheet = xlApp.Worksheets.Add(After: xlApp.ActiveSheet);

                string sheetname = Suffix + item;

                sheetname.Replace("\\","");
                sheetname.Replace("/", "");
                sheetname.Replace("*", "");
                sheetname.Replace("?", "");
                sheetname.Replace("[", "");
                sheetname.Replace("]", "");

                if (sheetname.Length > 31)
                {
                    sheetname = sheetname.Substring(0, 30);
                } 

                newWorksheet.Name = sheetname;
                from.Copy(newWorksheet.UsedRange);

                Result r = new Result { PageName = item, RowCount = ws.AutoFilter.Range.Columns[ColumnIndex].SpecialCells(XlCellType.xlCellTypeVisible).Cells.Count - (TableHeader - 1) };
                results.Add(r);

            }

            ws.EnableAutoFilter = false;

        }

        public bool CheckIntegrity()
        {
            int t = 0;
            int tPages = TotalRow - (TableHeader - 1);

            foreach (var item in results)
            {
                t = t + item.RowCount;
            }

            if (t == tPages)
            {
                return true;
            }

            return false;
        }

    }

    public class Result
    {
        public string PageName;
        public int RowCount;
    }
}
