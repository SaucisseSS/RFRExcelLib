using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace RFRExcelLib.Controllers
{
    public class TableController
    {
        private int _TableHeight { get; set; }
        private int _TableWidth { get; set; }
        private string _TablePosition { get; set; }
        public TableController() { }
        
        public TableController(int height, int width)
        {
            _TablePosition = "A1";
            _TableHeight = height;
            _TableWidth = width;
        }

        public void NewTable(string path)
        {

        }
    }
}
