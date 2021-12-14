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
        public string FolderLocation;
        public string Suffix;


        private int _ColumnIndex;
        public int ColumnIndex
        {
            get { return this._ColumnIndex; }
            set
            {
                if (this._ColumnIndex != value && value <= TotalCol && value > 0)
                    this._ColumnIndex = value;
            }
        }

        private int _TableHeader;
        public int TableHeader
        {
            get { return this._TableHeader; }
            set
            {
                if (this._TableHeader != value && value <= TotalRow && value > 0)
                    this._TableHeader = value;
            }
        }

        public HashSet<string> UniqueList;
        public int TotalRow;
        public int TotalCol;

        private Microsoft.Office.Interop.Excel.Application xlApp;
        private Microsoft.Office.Interop.Excel.Worksheet ws;

        public Export(Microsoft.Office.Interop.Excel.Application application)
        {
            FolderLocation = new KnownFolder(KnownFolderType.Downloads).Path;
            Suffix = DateTime.Today.ToString("yyyy.MM.dd") + "_";
            ColumnIndex = 1;
            TableHeader = 2;
            UniqueList = new HashSet<string>();

            xlApp = application;

            //Read Active sheet
            ws = xlApp.ActiveSheet;

            ws.EnableAutoFilter = true;

            TotalCol = ws.UsedRange.Columns.Count;
            TotalRow = ws.UsedRange.Rows.Count;

            UpdateUniqueList();
        }

        public void UpdateUniqueList()
        {
            object[,] array = ws.UsedRange.Value;

            for (int i = TableHeader; i < TotalRow; i++)
            {
                UniqueList.Add(array[i, ColumnIndex].ToString());
            }
        }

        public void ToNewFiles()
        {
            foreach (var item in UniqueList)
            {
                Range range = ws.UsedRange;
                range.AutoFilter(Field: ColumnIndex, Criteria1: item, VisibleDropDown: true);
                Range from = ws.UsedRange;

                try
                {

                    Workbook newbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                    Worksheet nws = newbook.Worksheets[1];
                    Range dest = nws.Range["A1"];

                    //Copy funziona solo se xlApp è la medesima!!!
                    from.Copy(dest);

                    newbook.SaveAs(FolderLocation + "\\" + Suffix + item);
                    newbook.Close();

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            }
        }

        public void ToNewSheets()
        {
            foreach (var item in UniqueList)
            {
                try
                {
                    Range range = ws.UsedRange;
                    range.AutoFilter(Field: ColumnIndex, Criteria1: item, VisibleDropDown: true);
                    Range from = ws.UsedRange;


                    Worksheet newWorksheet = xlApp.Worksheets.Add(After: xlApp.ActiveSheet);
                    from.Copy(newWorksheet.UsedRange);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            }
        }
    }
}
