using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JoinTable
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage ep = new ExcelPackage(new FileInfo("../../../../sensor_Kansas.xlsx"));
            ExcelWorksheet ws = ep.Workbook.Worksheets["Sheet1"];
            //var list = new List<Tuple<string, string>>();
            List<table> table_kansas = new List<table>();
            List<table> table_all = new List<table>();
            for (int rw = 2; rw <= ws.Dimension.End.Row; rw++)
            {
                if (ws.Cells[rw, 3].Value != null)
                    table_kansas.Add(new table {
                        key = ws.Cells[rw, 3].Value.ToString(),
                        value = ""
                    });
            }

            //ep = new ExcelPackage(new FileInfo("../../../../sensor_Kansas.xlsx"));
            ExcelWorksheet ws_all = ep.Workbook.Worksheets["Sheet2"];
            for (int rw = 2; rw <= ws_all.Dimension.End.Row; rw++)
            {
                if (ws_all.Cells[rw, 2].Value != null)
                    table_all.Add(new table
                    {
                        key = ws_all.Cells[rw, 2].Value.ToString(),
                        value = ws_all.Cells[rw, 1].Value.ToString(),
                    });
            }

            for (int i = 0; i < table_kansas.Count; i++)
            {
                for (int j = 0; j < table_all.Count; j++)
                {
                    if (table_kansas[i].key == table_all[j].key)
                        if(table_kansas[i].value != "")
                            table_kansas[i].value = table_kansas[i].value + ", " + table_all[j].value;
                        else
                            table_kansas[i].value = table_all[j].value;
                }
                Console.WriteLine("{0} , {1}", table_kansas[i].key, table_kansas[i].value);
            }

            

            int a = 2;


        }
    }
}
