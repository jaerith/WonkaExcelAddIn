using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace WonkaExcelAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public Excel.Range GetActiveCell()
        {
            return (Excel.Range) Application.ActiveCell;
        }

        public Dictionary<string,string> GetCurrentAttributeData()
        {
            var currData = new Dictionary<string, string>();

            var sheets = Application.ThisWorkbook.Worksheets;

            var worksheet = sheets.get_Item(1);

            /*
            for (int i = 1; i <= worksheet.Columns.Count; i++)
            {
                Excel.Range range = worksheet.get_Range("A" + i.ToString(), "J" + i.ToString());

                string value = range.Cells.Value2.ToString();

                if (value.Contains("orderID"){
                    //This column is for orders, I need to stop here and get all cell values under this column. 
                    break;
                }
            }
            */

            for (int i = 1; i <= worksheet.Rows.Count; i++)
            {
                Excel.Range range = worksheet.get_Range("A" + i.ToString(), "B" + i.ToString());

                string sAttrName  = range.Cells.Value.ToString();
                string sAttrValue = range.Cells.Value2.ToString();

                if (!String.IsNullOrEmpty(sAttrName) && !String.IsNullOrEmpty(sAttrValue))
                {
                    currData[sAttrName] = sAttrValue;
                }
            }

            return currData;
        }

        void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            /**
             ** NOTE: Not used for now
             ** 
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

            Excel.Range firstRow = activeWorksheet.get_Range("A1");

            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            Excel.Range newFirstRow = activeWorksheet.get_Range("A1");

            newFirstRow.Value2 = "This text was added by using code";
             **/
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup  += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
