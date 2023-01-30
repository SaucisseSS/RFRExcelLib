using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace RFRExelLib.Controllers
{
    public sealed class ExcelFileController
    {
        private Excel.Application _application;
        private Excel.Workbook _activeWorkbook;

        
        public string Workbook => _activeWorkbook.Name;
        public string Workbooks
        {
            get
            {
                string names = "";
                for (int i = 0; i < _application.Workbooks.Count; i++)
                    names += _application.Workbooks.Item[i].Name + '\n';
                return names;
            }
        }

        public ExcelFileController(string path)
        {
            _application = new Excel.Application();
        }
        public void AddWorkbook(string path)
        {
            if (path != null) _activeWorkbook = _application.Workbooks.Add(path);
            else _activeWorkbook = _application.Workbooks.Add(Missing.Value);
        }

        public bool FindWorkbook() => _application.FindFile();
        public void CloseWorkbook(string filename)
        {
            if (filename != null) _application.ActiveWorkbook.Close();
            else _application.ActiveWorkbook.Close(filename);
        }
       
        ~ExcelFileController()
        {
            _application.Workbooks.Close();
            _application.Quit();
        }
        
    }
}
