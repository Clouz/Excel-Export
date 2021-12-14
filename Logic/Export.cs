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

        public HashSet<string> UniqueList = new HashSet<string>();
        public int TotalRow;
        public int TotalCol;

        private Microsoft.Office.Interop.Excel.Application xlApp;
        private Microsoft.Office.Interop.Excel.Worksheet ws;
        private object[,] array;


        public Export(Microsoft.Office.Interop.Excel.Application application)
        {
            FolderLocation = new KnownFolder(KnownFolderType.Downloads).Path;
            Suffix = DateTime.Today.ToString("yyyy.MM.dd") + "_";
            ColumnIndex = 1;
            TableHeader = 2;

            xlApp = application;

            //Read Active sheet
            ws = xlApp.ActiveSheet as Worksheet;

            ws.EnableAutoFilter = true;

            TotalCol = ws.UsedRange.Columns.Count;
            TotalRow = ws.UsedRange.Rows.Count;

            array = ws.UsedRange.Value;

            UpdateUniqueList();
        }

        public void UpdateUniqueList()
        {

            for (int i = TableHeader; i < TotalRow; i++)
            {
                string part = array[i, ColumnIndex].ToString();
                UniqueList.Add(part);
            }
        }

        public void ToNewFiles()
        {
            foreach (var item in UniqueList)
            {
                Range range = ws.UsedRange;
                range.AutoFilter(Field: ColumnIndex, Criteria1: item, VisibleDropDown: true);
                Range from = ws.UsedRange;

                Workbook newbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                Worksheet nws = newbook.Worksheets[1];
                Range dest = nws.Range["A1"];

                //Copy funziona solo se xlApp è la medesima!!!
                from.Copy(dest);

                newbook.SaveAs(FolderLocation + "\\" + Suffix + item);
                newbook.Close();
            }
        }

        public void ToNewSheets()
        {
            foreach (var item in UniqueList)
            {
                Range range = ws.UsedRange;
                range.AutoFilter(Field: ColumnIndex, Criteria1: item, VisibleDropDown: true);
                Range from = ws.UsedRange;

                Worksheet newWorksheet = xlApp.Worksheets.Add(After: xlApp.ActiveSheet);
                from.Copy(newWorksheet.UsedRange);
            }
        }
    }
}
