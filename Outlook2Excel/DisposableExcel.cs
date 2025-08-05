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
        private string _fileName;
        public bool IsProgramInitiatedClose = false;
        public Application App => _excelApp;
        public Workbook Workbook => _workbook;
        public Worksheet Worksheet => _worksheet;
        private bool islocked = false;
        private bool _IsLocked
        {
            get
            {
                return islocked;
            }
            set
            {
                if(_excelApp != null)
                {
                    _excelApp.Interactive = !value;
                    _excelApp.ScreenUpdating = !value;
                }
            }
        }


        public List<string> PrimaryKeyValsAlreadyInExcel;
        public Dictionary<string, int> ExcelHeaders;
        private int _PrimaryKeyCol = 0;

        public DisposableExcel(string path)
        {
            _fileName = path;
            _CreateExcel(path);
        }
        private void _CreateExcel(string path)
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
            if (_worksheet == null) _worksheet = new Worksheet();

            _excelApp.WorkbookBeforeClose += IsUserTryingToCloseWB;
        }

        

        private void IsUserTryingToCloseWB(Workbook Wb, ref bool Cancel)
        {
            Cancel = !IsProgramInitiatedClose;
        }

        public void SaveAndClose()
        {
            try
            {
                IsProgramInitiatedClose = true;
                _workbook.Save();
                _workbook.Close(false);
                _excelApp.Quit();
            }
            catch(Exception ex)
            {
                StaticMethods.Quit("Excel failed to close properly. Make sure it is not open on server before trying again.", 203, ex);
            }            
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
                AppLogger.Log.Info("Locking Excel to perform data deposit");
                _IsLocked = true;
                AppLogger.Log.Info("Excel locked");
            }
            catch (Exception ex)
            {
                AppLogger.Log.Warn("Excel was being edited while cells were trying to be inserted. Aborting upload and trying again after timer.", ex);
                _IsLocked = false;
                return;
            }

            try
            {
                if (emailData == null || emailData.Count == 0)
                {
                    _IsLocked = false;
                    return;
                }

                if (ExcelHeaders.Count == 0)
                    GetOrSetExcelHeaders(emailData[0].Keys.ToArray());

                if (!string.IsNullOrEmpty(primaryKey))
                    PrimaryKeyValsAlreadyInExcel = GetPrimaryKeyValsInExcel(_worksheet, primaryKey).Distinct().ToList();

                int totalRows = emailData.Count;
                int startRow = GetLastRow(_PrimaryKeyCol) + 1;
                if (startRow <= 0) return;

                // Filter out rows with duplicate primary keys
                var rowsToInsert = emailData
                    .Where(d => string.IsNullOrEmpty(primaryKey) || !PrimaryKeyValsAlreadyInExcel.Contains(d[primaryKey]))
                    .ToList();

                int rowCount = rowsToInsert.Count;
                int colCount = ExcelHeaders.Count;

                if (rowCount == 0)
                {
                    _IsLocked = false;
                    return;
                }

                object[,] dataArray = new object[rowCount, colCount];

                for (int r = 0; r < rowCount; r++)
                {
                    var dict = rowsToInsert[r];
                    foreach (var key in dict.Keys)
                    {
                        if (ExcelHeaders.TryGetValue(key, out int colIndex))
                        {
                            //Excel is 1-based, array is 0-based
                            dataArray[r, colIndex - 1] = dict[key];
                        }
                    }
                    _excelApp.StatusBar = $"PROCESSING ROW {r + 1} of {rowCount} - {(int)((r + 1) / (double)rowCount * 100)}%";
                }

                //Write to worksheet in one operation
                var targetRange = _worksheet.Range[
                    _worksheet.Cells[startRow, 1],
                    _worksheet.Cells[startRow + rowCount - 1, colCount]
                ];
                targetRange.Value2 = dataArray;
            }
            catch (Exception ex)
            {
                StaticMethods.Quit("Generic Excel Error after load", 301, ex);
            }
            finally
            {
                _excelApp.StatusBar = "PROCESSING DONE";
                _IsLocked = false;
            }


        }
        private List<string> GetPrimaryKeyValsInExcel(Worksheet ws, string primaryKey)
        {
            List<string> colValues = new List<string>();

            if (primaryKey == "") return new List<string>();
            if (!ExcelHeaders.Keys.Contains(primaryKey)) return new List<string>();
            _PrimaryKeyCol = ExcelHeaders[primaryKey];

            //Get all range in PrimaryKey column
            int bottomRow = GetLastRow(_PrimaryKeyCol);
            if(bottomRow <= 0) return new List<string>();
            Microsoft.Office.Interop.Excel.Range primaryKeyRange = ws.Columns[_PrimaryKeyCol];
            object[,]? values = primaryKeyRange.Value2 as object[,];
            if(values == null) return new List<string>();

            //Get all values in PrimaryKey column
            for(int row = 2; row < bottomRow+1; row++)
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
            int lastRow;
            try
            {
                Microsoft.Office.Interop.Excel.Range lastCell = _worksheet.Cells[_worksheet.Rows.Count, column];
                Microsoft.Office.Interop.Excel.Range lastUsed = lastCell.End[XlDirection.xlUp];

                lastRow = lastUsed.Row;
                AppLogger.Log.Info($"Last used row found at {lastUsed.Row}.");
            }
            catch(Exception ex)
            {
                AppLogger.Log.Warn("Excel was busy when trying to find last row. Aborting row write. Trying again after timer");
                lastRow = -1;
            }
            // Avoid writing over data
            return lastRow;
        }
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void ResetExcelApp()
        {
            try
            {
                SaveAndClose();
                Marshal.FinalReleaseComObject(_excelApp);
                DisposeObject(_excelApp);
                DisposeObject(_worksheet);
                DisposeObject(_workbook);
                _CreateExcel(_fileName);
            }
            catch (Exception ex)
            {
                StaticMethods.Quit("Excel failed to restart after being restarted.", 500, ex);
            }
        }

        protected void DisposeObject(object? o)
        {
            if (o != null)
            {
                Marshal.ReleaseComObject(o);
                o = null;
            }
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
