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
            bool joinByName;
            bool joinByRef;
            ExcelPackage ep = new ExcelPackage(new FileInfo("../../../../sensor_Kansas.xlsx"));
            ExcelWorksheet ws = ep.Workbook.Worksheets["Sheet1"];
            //var list = new List<Tuple<string, string>>();
            List<table> table_kansas = new List<table>();
            List<table> table_all = new List<table>();
            List<table> joinName = new List<table>();
            for (int rw = 2; rw <= ws.Dimension.End.Row; rw++)
            {
                if (ws.Cells[rw, 3].Value != null)
                    table_kansas.Add(new table {
                        key = ws.Cells[rw, 3].Value.ToString(),
                        locationRef = ""
                    });
            }

            joinByRef = false;
            if (joinByRef)
            {
                //ep = new ExcelPackage(new FileInfo("../../../../sensor_Kansas.xlsx"));
                ExcelWorksheet ws_all = ep.Workbook.Worksheets["Sheet2"];
                for (int rw = 2; rw <= ws_all.Dimension.End.Row; rw++)
                {
                    if (ws_all.Cells[rw, 2].Value != null)
                        table_all.Add(new table
                        {
                            key = ws_all.Cells[rw, 2].Value.ToString(),
                            locationRef = ws_all.Cells[rw, 1].Value.ToString(),
                        });
                }

                for (int i = 0; i < table_kansas.Count; i++)
                {
                    for (int j = 0; j < table_all.Count; j++)
                    {
                        if (table_kansas[i].key == table_all[j].key)
                            if (table_kansas[i].locationRef != "")
                                table_kansas[i].locationRef = table_kansas[i].locationRef + ", " + table_all[j].locationRef;
                            else
                                table_kansas[i].locationRef = table_all[j].locationRef;
                    }
                    Console.WriteLine("{0} , {1}", table_kansas[i].key, table_kansas[i].locationRef);
                }
            }

            joinByName = true;

            if (joinByName)
            {
                ExcelWorksheet ws_all = ep.Workbook.Worksheets["Sheet2"];
                for (int rw = 2; rw <= ws_all.Dimension.End.Row; rw++)
                {
                    if (ws_all.Cells[rw, 3].Value != null)
                        table_all.Add(new table
                        {
                            key = ws_all.Cells[rw, 3].Value.ToString(),
                            locationRef = ws_all.Cells[rw, 2].Value.ToString(),
                            intId = ws_all.Cells[rw, 1].Value.ToString()
                        });
                }

                for (int i = 0; i < table_all.Count; i++)
                {
                    string key = table_all[i].key;
                    string locRef = table_all[i].locationRef;
                    string id = table_all[i].intId;
                    for (int j = i + 1; j < table_all.Count; j++)
                    {
                        if (table_all[i].key.Replace(" ","").ToLower().ToString() == table_all[j].key.Replace(" ","").ToLower().ToString())
                            if (table_all[i].locationRef != "")
                            {
                                locRef = locRef + "; " + table_all[j].locationRef;
                                id = id + "; " + table_all[j].intId;
                            }
                    }
                    joinName.Add(new table
                    {
                        key = key,
                        locationRef = locRef,
                        intId = id
                    });
                }

                foreach (table t in joinName)
                {
                    Console.WriteLine("{0} , {1} , {2}", t.key, t.locationRef, t.intId);
                }
            }

            int a = 2;
            Console.Read();

        }


    }
}
