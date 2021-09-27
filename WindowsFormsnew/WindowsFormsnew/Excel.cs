﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsnew
{
    class Excel
    {
        string path = "";

        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(String path, int sheet)
        {
            this.path = path;
            
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
            
        }


        public string ReadCell(int i, int j)
        {
            //i++;
            //j++;
            if (ws.Cells[i, j].value2 != null)
                return (ws.Cells[i, j].Value2).ToString();
            else
                return " ";

            excel.Workbooks.Close();
            excel.Quit();


        }

        
    }
}
