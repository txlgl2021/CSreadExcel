using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Console;
using Spire.Xls;
using Spire.Pdf;

namespace ReadExcelLIb
{
    class Program
    {
        
        static void Main(string[] args)
        {
            Workbook wb = new Workbook();
            wb.LoadFromFile(@"d:\lx\4500.xls");
            Worksheet sheet = wb.Worksheets[0];
            int a = sheet.LastRow;
            int b = sheet.LastColumn;
            WriteLine($"行{a},列{b}");
            foreach(var r in sheet.Rows)
            {
                foreach(var c in r)
                {
                    Write($"{c.Value},");
                }
                WriteLine();
            }
            //wb.SaveToFile(@"d:\lx\4500.pdf", Spire.Xls.FileFormat.PDF);

            ReadKey();


        }
    }
}
