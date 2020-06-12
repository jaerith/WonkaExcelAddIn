using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;

namespace WonkaExcelAddIn
{
    public partial class ThisAddIn
    {

        public Excel.Range GetActiveCell()
        {
            return (Excel.Range) Application.ActiveCell;
        }

        public Dictionary<string,string> GetCurrentAttributeData()
        {
            var currData = new Dictionary<string, string>();

            try
            {
                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

                for (int i = 1; i <= activeWorksheet.Rows.Count; i++)
                {
                    // Excel.Range currRow = activeWorksheet.get_Range("A" + i.ToString(), "B" + i.ToString());
                   
                    string sAttrName  = (string)(activeWorksheet.Cells[i, 1] as Excel.Range).Value;
                    string sAttrValue = "";

                    var attrVal = (activeWorksheet.Cells[i, 2] as Excel.Range).Value;
                    if (attrVal is String)
                        sAttrValue = (string) attrVal;
                    else if (attrVal is DateTime)
                    {
                        sAttrValue = Convert.ToString(attrVal);
                        if (sAttrValue.Contains(" "))
                        {
                            string[] DateParts = sAttrValue.Split(' ');
                            if (DateParts.Length > 0)
                                sAttrValue = DateParts[0];
                        }
                    }
                    else
                        sAttrValue = Convert.ToString(attrVal);

                    if (sAttrName == null)
                        break;

                    if (!String.IsNullOrEmpty(sAttrName) && !String.IsNullOrEmpty(sAttrValue))
                    {
                        currData[sAttrName] = sAttrValue;
                    }
                }
            }
            catch (Exception ex)
            {
                // 0x800A03EC?
                MessageBox.Show("ERROR!  Exception throw: (" + ex.Message + ")");
            }

            return currData;
        }

        public void SetCurrentAttributeData(Dictionary<string, string> poCurrData)
        {
            try
            {
                Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);

                List<string> keyList = poCurrData.Keys.ToList();

                for (int i = 1; i <= keyList.Count; i++)
                {
                    string sAttrName  = keyList[i-1];
                    string sAttrValue = poCurrData[sAttrName];

                    (activeWorksheet.Cells[i, 1] as Excel.Range).Value = sAttrName;
                    (activeWorksheet.Cells[i, 2] as Excel.Range).Value = sAttrValue;
                }
            }
            catch (Exception ex)
            {
                // 0x800A03EC
                MessageBox.Show("ERROR!  Exception throw: (" + ex.Message + ")");
            }
        }

        #region Handlers 

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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

        #endregion

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
