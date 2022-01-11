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
              <button id='ExportToFiles' label='Export to Files' onAction='OnButtonPressedExportToFiles'/>
              <button id='ExportToSheets' label='Export to Sheets' onAction='OnButtonPressedExportToSheets'/>
            </group >
            <group id='group2' label='About'>
              <button id='About' label='About' onAction='OnButtonPressedAbout'/>
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
                Finestra.MainWindow win = new Finestra.MainWindow((Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application);
                win.ShowDialog();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "Error");
            }
        }

        public void OnButtonPressedExportToSheets(IRibbonControl control)
        {
            try
            {
                Finestra.MainWindow win = new Finestra.MainWindow((Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application, false);
                win.ShowDialog();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "Error");
            }
        }

        public void OnButtonPressedAbout(IRibbonControl control)
        {
            MessageBox.Show("Version 0.2\nCopyright(C) 2021 by Claudio Mola\n\nMore information: https://github.com/Clouz/Excel-Export \n\nThis program is free software: you can redistribute it and / or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.\n\nThis program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.\nYou should have received a copy of the GNU General Public License along with this program. If not, see < http://www.gnu.org/licenses/>.\n\n", "About");
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