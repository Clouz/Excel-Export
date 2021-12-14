using System.Runtime.InteropServices;
using System.Windows;
using ExcelDna.Integration.CustomUI;
using Excel_Export;
using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System.Collections.Generic;
using System;

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
          <tab id='tab1' label='My Tab'>
            <group id='group1' label='My Group'>
              <button id='button1' label='My Button' onAction='OnButtonPressed'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            //Finestra f = new Finestra();
            //f.UpdateText("Ciao from control " + control.Id);

            Microsoft.Office.Interop.Excel.Application xlApp = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;

            Worksheet ws = xlApp.ActiveSheet as Worksheet;

            //Add data into A1 and B1 Cells as headers.
            ws.Cells[1, 1] = "Product ID";
            ws.Cells[1, 2] = "Product Name";

            //Add data into details cells.
            ws.Cells[2, 1] = "aa";
            ws.Cells[3, 1] = "aa";
            ws.Cells[4, 1] = "bb";
            ws.Cells[5, 1] = "aa";
            ws.Cells[6, 1] = "Bb";
            ws.Cells[2, 2] = "Apples";
            ws.Cells[3, 2] = "Bananas";
            ws.Cells[4, 2] = "Grapes";
            ws.Cells[5, 2] = "Oranges";
            ws.Cells[6, 2] = "Raspberry";


            ws.EnableAutoFilter = true;

            int colNo = ws.UsedRange.Columns.Count;
            int rowNo = ws.UsedRange.Rows.Count;
            object[,] array = ws.UsedRange.Value;
            HashSet<string> uniqueList = new HashSet<string>();
            int col = 1;

            for (int i = 2; i < rowNo; i++)
            {
                uniqueList.Add(array[i, col].ToString());
            }

            foreach (var item in uniqueList)
            {
                Range range = ws.UsedRange;
                range.AutoFilter(Field: col, Criteria1: item, VisibleDropDown: true);
                Range from = ws.UsedRange;


                //Worksheet newWorksheet = xlApp.Worksheets.Add(After: xlApp.ActiveSheet);
                //from.Copy(newWorksheet.UsedRange);

                try
                {

                    string FileDropLocation = @"C:\users-data\SACOA002036\Downloads\Nuova cartella (3)";
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    app.Visible = false;
                    Workbook newbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                    Worksheet nws = newbook.Worksheets[1];
                    Range dest = nws.Range["A1"];
                    from.Copy(dest);


                    //from.Copy(newbook.Worksheets[1]);

                    newbook.SaveAs(FileDropLocation + "\\" + "asd_" + item);
                    newbook.Close();
                    app.Quit();


                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }



            }

            ws.AutoFilterMode = false;


        }

        private void saveInNewExcel(Worksheet from, string name) {

            try
            {

                string FileDropLocation = @"C:\users-data\SACOA002036\Downloads\Nuova cartella (3)";
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = false;
                Workbook newbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                Worksheet newWorksheet = newbook.Worksheets[1];
                newWorksheet.Cells[1, 1] = "prova";

                //from.Copy(newbook.Worksheets[1]);

                newbook.SaveAs(FileDropLocation + "\\" + "asd_" +name);
                newbook.Close();
                app.Quit();
               

            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }

        }
    }
}