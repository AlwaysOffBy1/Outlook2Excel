using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Outlook2Excel.Core;

namespace Outlook2Excel
{
    public class DisposableExcel : IDisposable
    {
        private Application _excelApp;
        private Workbook _workbook;
        private Worksheet _worksheet;
        private bool _disposed = false;
        public bool IsProgramInitiatedClose = false;
        public Application App => _excelApp;
        public Workbook Workbook => _workbook;
        public Worksheet Worksheet => _worksheet;


        public List<string> PrimaryKeyValsAlreadyInExcel;
        public Dictionary<string, int> ExcelHeaders;

        public DisposableExcel(string path)
        {
            if (!File.Exists(path))
                StaticMethods.Quit("Excel file path not found", 100, null);

            PrimaryKeyValsAlreadyInExcel = new List<string>();
            ExcelHeaders = new Dictionary<string, int>();

            _excelApp = new Application();
            _excelApp.Visible = true;

            try
            {
                _workbook = _excelApp.Workbooks.Open(path, ReadOnly: false, Notify: true);
                _worksheet = (Worksheet)_workbook.Sheets[1];
            }
            catch (COMException ex)
            {
                StaticMethods.Quit("Failed to open Excel file. It might already be open.\n" + ex.Message, 301, null);
            }

            if (_workbook == null) _workbook = new Workbook();
            if(_worksheet == null) _worksheet = new Worksheet();

            _excelApp.WorkbookBeforeClose += IsUserTryingToCloseWB;
        }

        private void IsUserTryingToCloseWB(Workbook Wb, ref bool Cancel)
        {
            Cancel = !IsProgramInitiatedClose;
        }

        public void SaveAndClose()
        {
            IsProgramInitiatedClose = true;
            _workbook.Save();
            _workbook.Close(false);
            _excelApp.Quit();
        }

        private void GetOrSetExcelHeaders(string[] dataHeaders)
        {
            //Write headers if none exist already
            if (string.IsNullOrEmpty(_worksheet.Cells[1, 1].Value2))
                for (int col = 0; col < 100; col++)
                {
                    _worksheet.Cells[1, col + 1] = dataHeaders[col];
                    ExcelHeaders.Add(dataHeaders[col], col);
                    return;
                }

            //Otherwise, read headers
            for (int col = 0; col < 100; col++)
            {
                string header = _worksheet.Cells[1, col + 1].Value2;
                if (string.IsNullOrEmpty(header)) break;
                ExcelHeaders.Add(header, col + 1);
            }

            if (ExcelHeaders.Count < dataHeaders.Count()) Console.WriteLine("!!! Not all excel fields match up with your outlook fields. Some fields will be missing!");
            return;
        }
        public void AddData(List<Dictionary<string, string>> emailData, string primaryKey)
        {
            try
            {
                _excelApp.ScreenUpdating = false;
                _excelApp.Interactive = false;
            }
            catch(Exception ex)
            {
                AppLogger.Log.Warn("Excel was being edited while cells were trying to be inserted. Aborting upload and trying again after timer.");
                return;
            }
            

            //Dont like large try's, but user can interfere at any time in so many ways
            try
            {
                if (emailData == null || emailData.Count == 0)
                    return;

                if(ExcelHeaders.Count == 0) GetOrSetExcelHeaders(emailData[0].Keys.ToArray());
                //Shouldn't have 2 identical primary keys, so get list of all of them before writing
                //Since each excel file can be open by more than 1 person at a time, and can be edited by more
                //than one person at a time, this needs to be checked each time an edit wants to happen
                if (primaryKey != "") PrimaryKeyValsAlreadyInExcel = GetPrimaryKeyValsInExcel(_worksheet, primaryKey).Distinct().ToList();

                //Write each row of data
                double dictRows = emailData.Count;
                int startRow = GetLastRow() +1;
                for (int row = 0; row < dictRows; row++)
                {
                    var currentDict = emailData[row];

                    //If primary key is already in excel skip it
                    string val = currentDict[primaryKey];
                    if (PrimaryKeyValsAlreadyInExcel.Contains(currentDict[primaryKey])) continue;
                    _excelApp.StatusBar = $"PROCESSING ROW {row} of {dictRows} - {(int)((row/ dictRows) *100)}%";

                    int i = 1;
                    foreach (var key in currentDict.Keys)
                    {
                        if(ExcelHeaders.Keys.Contains(key))
                            _worksheet.Cells[startRow + row, ExcelHeaders[key]] = currentDict[key];
                    }
                }
            }
            catch(Exception ex)
            {
                StaticMethods.Quit("Generic Excel Error after load", 301, ex);
            }
            _excelApp.StatusBar = "PROCESSING DONE";
            _excelApp.ScreenUpdating = true;
            _excelApp.Interactive = true;
            
        }
        private List<string> GetPrimaryKeyValsInExcel(Worksheet ws, string primaryKey)
        {
            List<string> colValues = new List<string>();

            if (primaryKey == "") return new List<string>();
            int primaryKeyCol = 0;
            if (!ExcelHeaders.Keys.Contains(primaryKey)) return new List<string>();
            primaryKeyCol = ExcelHeaders[primaryKey];

            //Get all range in PrimaryKey column
            int bottomRow = GetLastRow(primaryKeyCol);
            Microsoft.Office.Interop.Excel.Range primaryKeyRange = ws.Columns[primaryKeyCol];
            object[,]? values = primaryKeyRange.Value2 as object[,];
            if(values == null) return new List<string>();

            //Get all values in PrimaryKey column
            for(int row = 2; row < bottomRow; row++)
            {
                object val = values[row,1];
                if(val != null)
                    colValues.Add(val.ToString() ?? "");
            }
            colValues = colValues.Distinct().ToList();
            return colValues;
        }
        private int GetLastRow(int column = 1)
        {
            Microsoft.Office.Interop.Excel.Range lastCell = _worksheet.Cells[_worksheet.Rows.Count, column];
            Microsoft.Office.Interop.Excel.Range lastUsed = lastCell.End[XlDirection.xlUp];

            AppLogger.Log.Info($"Last used row found at {lastUsed.Row}.");
            int lastRow = lastUsed.Row;

            // Avoid writing over data
            return lastRow;
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            // Release COM objects
            if (_worksheet != null) Marshal.ReleaseComObject(_worksheet);
            if (_workbook != null) Marshal.ReleaseComObject(_workbook);
            if (_excelApp != null) Marshal.ReleaseComObject(_excelApp);

            _worksheet = null;
            _workbook = null;
            _excelApp = null;

            _disposed = true;
        }

        ~DisposableExcel()
        {
            Dispose(false);
        }
    }
}
