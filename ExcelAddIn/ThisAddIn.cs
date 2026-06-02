using System;
using CommonLibrary;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private TaskPaneControl _taskPaneControl;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _taskPaneControl = new TaskPaneControl();
            _taskPane = CustomTaskPanes.Add(_taskPaneControl, "OfficePractice");
            _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionBottom;
            _taskPane.Height = 250;
            _taskPane.Visible = true;

            Application.SheetChange += OnSheetChange;
            Application.WorkbookOpen += OnWorkbookOpen;
            Application.WorkbookBeforeClose += OnWorkbookBeforeClose;
            Application.WorkbookActivate += OnWorkbookActivate;

            SendAndLog("AddInLoad", "Excel add-in loaded successfully");
        }

        private void OnSheetChange(object sh, Excel.Range target)
        {
            var sheetName = sh is Excel.Worksheet ws ? ws.Name : "?";
            SendAndLog("SheetChange", $"Sheet: {sheetName}, Cell: {target.Address}");
        }

        private void OnWorkbookOpen(Excel.Workbook wb)
        {
            SendAndLog("WorkbookOpen", $"File: {wb.Name}");
        }

        private void OnWorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            SendAndLog("WorkbookClose", $"File: {wb.Name}");
        }

        private void OnWorkbookActivate(Excel.Workbook wb)
        {
            SendAndLog("WorkbookActivate", $"File: {wb.Name}");
        }

        private void SendAndLog(string eventType, string detail)
        {
            var msg = new PipeMessage { EventType = eventType, Application = "Excel", Detail = detail };
            PipeClient.Send(msg);
            _taskPaneControl?.AddLog($"{eventType}: {detail}");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.SheetChange -= OnSheetChange;
            Application.WorkbookOpen -= OnWorkbookOpen;
            Application.WorkbookBeforeClose -= OnWorkbookBeforeClose;
            Application.WorkbookActivate -= OnWorkbookActivate;
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
