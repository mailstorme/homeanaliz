using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Подсчет_начислений
{
    class OpenExcelFile
    {
        public OpenFileDialog dlg = new OpenFileDialog();
        public string Filenamereturn()
        {
            dlg.FileName = "Файл Excel"; // Default file name
            //dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Excel (*.xlsx; *.xlsb)|*.xlsx; *.xlsb"; // Filter files by extension
            if (dlg.ShowDialog() == true)
            {
                // Open document
                return dlg.FileName;
            }
            else
                return "can not open file";
        }

    }
}
