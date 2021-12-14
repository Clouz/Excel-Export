using System.Runtime.InteropServices;
using System.Windows;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System.Collections.Generic;
using System;
using Finestra;

namespace Ribbon
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='tab1' label='Export'>
            <group id='group1' label='Export'>
              <button id='ExportToFiles' label='Export' onAction='OnButtonPressedExportToFiles'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressedExportToFiles(IRibbonControl control)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
                Excel_Export.Export data = new Excel_Export.Export(xlApp);

                MainWindow win = new MainWindow(data);
                win.ShowDialog();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
    }
}



//string FolderLocation = @"C:\users-data\SACOA002036\Downloads\Nuova cartella (3)";
//string Suffix = "2021.12.14_";
//int ColumnIndex = 1;

////Load current excel application
//Microsoft.Office.Interop.Excel.Application xlApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

////Read Active sheet
//Worksheet ws = xlApp.ActiveSheet as Worksheet;

//ws.EnableAutoFilter = true;

//int colNo = ws.UsedRange.Columns.Count;
//int rowNo = ws.UsedRange.Rows.Count;
//object[,] array = ws.UsedRange.Value;
//HashSet<string> uniqueList = new HashSet<string>();


//for (int i = 2; i < rowNo; i++)
//{
//    uniqueList.Add(array[i, ColumnIndex].ToString());
//}

//foreach (var item in uniqueList)
//{
//    Range range = ws.UsedRange;
//    range.AutoFilter(Field: ColumnIndex, Criteria1: item, VisibleDropDown: true);
//    Range from = ws.UsedRange;


//    //Worksheet newWorksheet = xlApp.Worksheets.Add(After: xlApp.ActiveSheet);
//    //from.Copy(newWorksheet.UsedRange);

//    try
//    {



//        Workbook newbook = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

//        Worksheet nws = newbook.Worksheets[1];
//        Range dest = nws.Range["A1"];

//        //Copy funziona solo se xlApp è la medesima!!!
//        from.Copy(dest);

//        newbook.SaveAs(FolderLocation + "\\" + Suffix + item);
//        newbook.Close();

//    }
//    catch (Exception e)
//    {
//        MessageBox.Show(e.ToString());
//    }

//}

//ws.AutoFilterMode = false;