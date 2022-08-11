using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IronXL;

namespace BollingerBand
{
    public class Utility
    {
        public static DataTable ReadExcel(string FileName)
        {
            WorkBook workbook = WorkBook.Load(FileName);
            WorkSheet sheet = workbook.DefaultWorkSheet;
            return sheet.ToDataTable(true);
        }

        public static void printDataTable(DataTable mytable)
        {
            var mycols = mytable.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();
            Console.WriteLine(string.Join(",", mycols));
            mycols.RemoveAt(0);

            foreach (DataRow row in mytable.Rows)
            {
                bool checkempty = true;
                foreach (string ordercolname in mycols)
                {
                    object value = row[ordercolname];
                    if (!(value == DBNull.Value))
                    {
                        checkempty = false;
                        break;
                    }
                }
                if (!checkempty)
                {
                    string printrow = string.Join(",", row.ItemArray.Select(p => p.ToString()).ToArray());
                    Console.WriteLine(printrow);
                }
            }
        }

        public static void csvDataTable(DataTable mytable, string filename)
        {
            var mycols = mytable.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();
            var checkcols = mycols.ToList();
            checkcols.RemoveAt(0);

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Join(",", mycols));

            foreach (DataRow row in mytable.Rows)
            {
                bool checkempty = true;
                foreach (string ordercolname in checkcols)
                {
                    object value = row[ordercolname];
                    if (!(value == DBNull.Value))
                    {
                        checkempty = false;
                        break;
                    }
                }
                if (!checkempty)
                {
                    string printrow = string.Join(",", row.ItemArray.Select(p => p.ToString()).ToArray());
                    sb.AppendLine(string.Join(",", printrow));
                }
            }
            File.WriteAllText(filename, sb.ToString());
        }
    }
}
