using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace xl2vsoAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            VSOConfig.Initialize();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            Converter _converter = new Converter(activeWorksheet);
            bool isSchemaValid = _converter.CheckColumnSchema();
            if (isSchemaValid)
            {
                VSOWorker _worker = new VSOWorker();
                char titleCol = 'C'; // using title column since title can not be blank
                int rowIndex = 2;
                while(true)
                {
                        string title = activeWorksheet.get_Range(titleCol + rowIndex.ToString()).Value2;
                        if(string.IsNullOrEmpty(title)) {
                            break; // stop
                        }
                        try
                        {
                            xl2vsoModel model = _converter.ConvertToVSOModel(rowIndex);
                            int id = _worker.CreateWorkItem(model);
                            //set success
                            activeWorksheet.get_Range('A' + rowIndex.ToString()).Value2 = "Success";
                            activeWorksheet.get_Range('B' + rowIndex.ToString()).Value2 = id.ToString();
                        }
                        catch (Exception e)
                        {
                            //set failure
                            activeWorksheet.get_Range('A' + rowIndex.ToString()).Value2 = "Fail : " + e.InnerException;
                        }
                        ++rowIndex;
                }
            }
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
