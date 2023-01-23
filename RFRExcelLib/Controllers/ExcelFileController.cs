using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Controllers.RFRExelLib
{
    public sealed class ExcelFileController
    {
        private Excel.Application _application;
        private List<Excel.Workbook> _workbooks;

        public List<Excel.Workbook> Workbooks => _workbooks;

        public ExcelFileController()
        {
            _application = new Excel.Application();
        }
        public ExcelFileController(string path)
        {
            _application = new Excel.Application();
            _workbooks.Add(_application.Workbooks.Add(path));
        }

        public void AddWorkBook(string path)
        {
            if (path != null) _workbooks.Add(_application.Workbooks.Add(path));
            else _workbooks.Add(_application.Workbooks.Add(Missing.Value));                   
        }



    }
}
