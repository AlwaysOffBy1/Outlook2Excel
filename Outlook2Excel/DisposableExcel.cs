using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading.Tasks;

namespace Outlook2Excel
{
    public class DisposableExcel : IDisposable
    {
        private Application _excelApp;
        private Workbook _workbook;
        private Worksheet _worksheet;
        private bool _disposed = false;

        public Application App => _excelApp;
        public Workbook Workbook => _workbook;
        public Worksheet Worksheet => _worksheet;

        public DisposableExcel()
        {
            _excelApp = new Application();
            _excelApp.Visible = true;
            _workbook = _excelApp.Workbooks.Add();
            _worksheet = (Worksheet)_workbook.Sheets[1];
        }

        public void SaveAndClose(string filePath)
        {
            _workbook.SaveAs(filePath);
            _workbook.Close(false);
            _excelApp.Quit();
        }


        public void AddData(List<Dictionary<string, string>> data)
        {
            if (data == null || data.Count == 0)
                return;

            var headers = data[0].Keys.ToList();

            // Write headers
            for (int col = 0; col < headers.Count; col++)
            {
                _worksheet.Cells[1, col + 1] = headers[col];
            }

            // Write each row of data
            for (int row = 0; row < data.Count; row++)
            {
                var dict = data[row];
                for (int col = 0; col < headers.Count; col++)
                {
                    try
                    {
                        if (headers.Contains(dict[headers[col]]))
                            _worksheet.Cells[row + 2, col + 1] = dict[headers[col]];
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"KEY ({dict[headers[col]]}) WAS NOT IN DICTIONARY ({dict[AppSettings.PrimaryKey]}");
                    }
                    
                }
            }
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
