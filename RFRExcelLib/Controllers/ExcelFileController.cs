using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace RFRExelLib.Controllers
{
    public sealed class ExcelFileController
    {
        private Excel.Application _application;
        private Excel.Workbook _workbooks;

        public string Workbook => _workbooks.FullName;

        public ExcelFileController()
        {
            _application = new Excel.Application();
            _workbooks = _application.Workbooks.Add(Missing.Value);
        }
        public ExcelFileController(string path)
        {
            _application = new Excel.Application();
            _workbooks = _application.Workbooks.Add(path);
        }
        public void AddWorkbook(string path)
        {
            if (path != null) _workbooks = _application.Workbooks.Add(path);
            else
            {
                var a = _application.Workbooks.Add(Missing.Value);
                _workbooks = a;
            }
        }
        public void CloseWorkbook()
        {

        }

        
    }
}
