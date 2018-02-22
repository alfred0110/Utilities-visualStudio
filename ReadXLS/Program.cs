using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadXLS
{
    class Program
    {
        static void Main(string[] args)
        {

            string[] files = Directory.GetFiles(@"C:\Users\Alfred\Desktop\FTP FIDEGAR\");

            for (int i = 0; i< files.Count(); i++ )
            {               
                var total =  Read_From_Excel.getExcelFile(files[i]);


                FileStream ostrm;
                StreamWriter writer;
                TextWriter oldOut = Console.Out;
                try
                {
                    ostrm = new FileStream("./Redirect.txt", FileMode.Append, FileAccess.Write);
                    writer = new StreamWriter(ostrm);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Cannot open Redirect.txt for writing");
                    Console.WriteLine(e.Message);
                    return;
                }
                Console.SetOut(writer);
                Console.WriteLine(string.Format("{0,6} {1,15} {2,30}", i + 1, total, files[i]));
                Console.SetOut(oldOut);
                writer.Close();
                ostrm.Close();
                Console.WriteLine(string.Format("{0,6} {1,15} {2,30}", i + 1, total, files[i]));
            }
            
        }
    }


    public class Read_From_Excel
    {
        public static double getExcelFile(string file)
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            double sum = 0;
            try
            {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;



                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!

                for (int i = 2; i <= rowCount; i++)
                {

                    sum += double.Parse(xlRange.Cells[i, 2].Value2.ToString());


                    //for (int j = 1; j <= colCount; j++)
                    //{
                    //    //new line
                    //    if (j == 1)
                    //        Console.Write("\r\n");

                    //    //write the value to the console
                    //    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    //        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                    //}
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (COMException ex)
            {
                return -1;
            }
            return sum;
        }
    }
}

