using ConsoleApp1;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace FastExceldemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Get the input file path
            var inputFile = new FileInfo("D:\\New Folder\\App.xlsx");

            int i, j;

            // Create an instance of Fast Excel
            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, true))
            {
                foreach (var worksheet in fastExcel.Worksheets)
                {
                    Console.WriteLine(string.Format("Worksheet Name:{0}, Index:{1}", worksheet.Name, worksheet.Index));

                    //To read the rows call read
                    worksheet.Read();
                    var rows = worksheet.Rows.ToArray();
                    //Do something with rows
                    Console.WriteLine(string.Format("Worksheet Rows:{0}", rows.Count()));
                  

                }
            }

     
        }

      
    }
}
