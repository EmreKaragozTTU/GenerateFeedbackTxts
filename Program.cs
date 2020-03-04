using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = "C:/Users/Emre-USA/Desktop/Homework Grading/Hw1Grading.xlsx";
            FileInfo fileInfo = new FileInfo(path);
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
            int rows = worksheet.Dimension.Rows;
            int columns = worksheet.Dimension.Columns;
            for (int i = 2; i <= rows; i++)
            {  
                string studentname = worksheet.Cells[i, 2].Value.ToString();
                using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:/Users/Emre-USA/Desktop/Homework Grading/Comments/"+ studentname + ".txt", true))
                for (int j = 3; j <= columns; j++)
                {
                    try {
                        string theComment = worksheet.Cells[1, j].Value.ToString();
                        string deduction = worksheet.Cells[i, j].Value.ToString();
                        string combination = theComment + ": " + deduction + "\n";
                        file.WriteLine(combination);
                        //Console.Write(cnt);
                        //Console.Write("\t");
                    }         
                    catch (Exception e) {
                        //Console.Write("\t");
                    }
                }
            }
            Console.ReadKey();
        }
    }
}
