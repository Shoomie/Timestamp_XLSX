using Microsoft.Office.Interop.Excel;
using System.IO;

namespace Timestamp_XLSX
{
    internal class Program
    {
        private static void ElapsedTime()
        {
            ConsoleKeyInfo keyInfo;
            DateTime startTime = DateTime.Now;

            Console.WindowHeight = 6;
            Console.WindowWidth = 60;

            while (true)
            {
                Console.Clear();

                TimeSpan elapsedTime = DateTime.Now - startTime;
                Console.WriteLine("Elapsed time: {0:00}:{1:00}:{2:00}", elapsedTime.Hours, elapsedTime.Minutes, elapsedTime.Seconds);
                Console.WriteLine("\nPress P to take a break and mark it in the time card,\nremember to start the counter again when you are back!");

                if (Console.KeyAvailable)
                {
                    keyInfo = Console.ReadKey(true);
                    if (keyInfo.Key == ConsoleKey.P)
                    {
                        break;
                    }
                }

                Thread.Sleep(1000);
            }
        }
        static void Main(string[] args)
        {
            string currentDirectory = Directory.GetCurrentDirectory();

            string currentYearAndMonth = DateTime.Now.ToString("yyyy-MM");
            string fileName = $"{currentYearAndMonth}.xlsx";
            Application excel = new Application();

            if (File.Exists(Path.Combine(currentDirectory, fileName)))
            {
                Workbook workbook = excel.Workbooks.Open(Path.Combine(currentDirectory, fileName));
                Worksheet worksheet = workbook.ActiveSheet;
                excel.Visible = false;
                int currentRow = -1;
                Microsoft.Office.Interop.Excel.Range range = worksheet.Range["A1", "A1000000"]; // range of cells in column A
                for (int row = 1; row <= range.Rows.Count; row++)
                {
                    if (range[row, 1].Value2 == null) // check if the cell is empty
                    {
                        currentRow = row;
                        worksheet.Cells[row, "A"] = DateTime.Now;
                        break;
                    }
                }
                if (currentRow != -1)
                {
                    ElapsedTime();

                    worksheet.Cells[currentRow, "B"] = DateTime.Now;
                    Microsoft.Office.Interop.Excel.Range cellA = worksheet.Cells[currentRow, 1];
                    Microsoft.Office.Interop.Excel.Range cellB = worksheet.Cells[currentRow, 2];
                    DateTime dateTimeA = DateTime.FromOADate((double)cellA.Value2);
                    DateTime dateTimeB = DateTime.FromOADate((double)cellB.Value2);
                    TimeSpan timeDifference = dateTimeB - dateTimeA;
                    worksheet.Cells[currentRow, "C"] = timeDifference.ToString();
                    workbook.Save();
                    workbook.Close();
                }
            }
            else
            {

                excel.Visible = false;
                Workbook workbook = excel.Workbooks.Add();
                Worksheet worksheet = workbook.ActiveSheet;
                worksheet.Cells[1, "A"] = "Start time";
                worksheet.Cells[1, "B"] = "End time";
                worksheet.Cells[1, "C"] = "Total Time";
                worksheet.Cells[2, "A"] = DateTime.Now;

                ElapsedTime();

                worksheet.Cells[2, "B"] = DateTime.Now;
                Microsoft.Office.Interop.Excel.Range cellA = worksheet.Cells[2, 1];
                Microsoft.Office.Interop.Excel.Range cellB = worksheet.Cells[2, 2];
                DateTime dateTimeA = DateTime.FromOADate((double)cellA.Value2);
                DateTime dateTimeB = DateTime.FromOADate((double)cellB.Value2);
                TimeSpan timeDifference = dateTimeB - dateTimeA;
                worksheet.Cells[2, "C"] = timeDifference.ToString();
                worksheet.Columns.AutoFit();
                workbook.SaveAs(Path.Combine(currentDirectory, fileName));
                workbook.Close();
            }
            excel.Quit();
            Environment.Exit(0);
        }
    }
}