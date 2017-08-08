using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace KeepItOpen
{
    public class fileHandler
    {
        // ------------------------------------ Declarations ------------------------------------
        public string openMode      = "";
        public string filePath      = "";
        public string filePathExcel = "";
        public int fileFormat       = -1;        // 0: word , 1: excel , -1: other
        public int timeToKeepOpen1  = 0;
        public int timeToKeepOpen2  = 0;
        public bool silentMode      = false;
        
        // ~~ lorem*X*: X = number of Lorem Ipsum paragraphs
        public string[] lorems = { loremsClass.lorem1, loremsClass.lorem5, loremsClass.lorem10, loremsClass.lorem15, loremsClass.lorem20 };
        // ---------------------------------------------------------------------------------------
        
        // ----------------------------- Open & handle word/excel files --------------------------
        public void openFile()
        {
            bool readOnly;
            if (openMode == "w")
                readOnly = false;
            else
                readOnly = true;

            switch (fileFormat)
            {
                case 0:
                    try
                    {
                        // ~~~~~~~~~~~~~~~~~~~~~~ Starting Word instance ~~~~~~~~~~~~~~~~~~~~
                        var wordApp = new Microsoft.Office.Interop.Word.Application();
                        wordApp.ShowAnimation = false;
                        wordApp.Visible = false;
                        object wordAppMissing = System.Reflection.Missing.Value;
                        Microsoft.Office.Interop.Word.Document wordFile = new Microsoft.Office.Interop.Word.Document();
                        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        // ~~ Opening Word file
                        wordFile = wordApp.Documents.Open(filePath, ref wordAppMissing, readOnly, ref wordAppMissing, 
                                                            ref wordAppMissing, ref wordAppMissing, ref wordAppMissing, 
                                                            ref wordAppMissing, ref wordAppMissing, ref wordAppMissing, 
                                                            ref wordAppMissing, false, ref wordAppMissing, ref wordAppMissing, 
                                                            ref wordAppMissing, ref wordAppMissing);
                        
                        // ~~ Read mode: delay and then close file
                        if (readOnly == true)
                        {
                            delayFunc(timeToKeepOpen1);

                            wordFile.Close(false, ref wordAppMissing, ref wordAppMissing);
                            wordFile = null;
                        }
                        // ~~ Write mode: delay1 - write data - ?delay2? - close file
                        else
                        {
                            // Get random Lorem Ipsum text to write
                            Random r = new Random(DateTime.Now.Millisecond);
                            int randomLorem = r.Next(0, 5);

                            // first delay
                            delayFunc(timeToKeepOpen1);
                            
                            wordFile.Content.SetRange(0, 0);
                            wordFile.Content.Text = lorems[randomLorem] + "\r\n";

                            // second delay (if any)
                            delayFunc(timeToKeepOpen2);

                            // saving & closing file
                            wordFile.SaveAs2(filePath);
                            wordFile.Close(false, ref wordAppMissing, ref wordAppMissing);
                            wordFile = null;
                        }

                        // ~~~~~~~~~~~~~~~~~~~~ Terminating Word instance ~~~~~~~~~~~~~~~~~~~
                        if (wordFile != null) Marshal.ReleaseComObject(wordFile);
                        wordApp.Quit(ref wordAppMissing, ref wordAppMissing, ref wordAppMissing);
                        wordApp.Quit();
                        if (wordApp != null) Marshal.ReleaseComObject(wordApp);
                        wordApp = null;
                        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error while opening word file");
                        Console.WriteLine(ex.Message);
                    }
                    break;
                case 1:
                    try
                    {
                        // ~~~~~~~~~~~~~~~~~~~~~ Starting Excel instance ~~~~~~~~~~~~~~~~~~~~
                        Microsoft.Office.Interop.Excel.Application excelApp;
                        Microsoft.Office.Interop.Excel.Workbook workBook;
                        Microsoft.Office.Interop.Excel.Worksheet workSheet;
                        excelApp = new Microsoft.Office.Interop.Excel.Application();
                        object excelAppMissing = System.Reflection.Missing.Value;
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;
                        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        // ~~ Opening Excel file
                        workBook = excelApp.Workbooks.Open(filePath, excelAppMissing, readOnly);

                        // ~~ Read mode: delay and then close file
                        if (readOnly == true)
                        {
                            delayFunc(timeToKeepOpen1);

                            workBook.Close(false);
                            workBook = null;
                        }
                        // ~~ Write mode: delay1 - write data - ?delay2? - close file
                        else
                        {
                            // Selecting random rows and columns count
                            Random r = new Random(DateTime.Now.Millisecond);
                            int rowsToWrite = r.Next(10, 70);
                            int colsToWrite = r.Next(10, 30);

                            // Opening & naming worksheet
                            workSheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                            workSheet.Name = "Edited file";

                            // first delay
                            delayFunc(timeToKeepOpen1);

                            // writting random numbers
                            var data = new object[rowsToWrite, colsToWrite];
                            for (var row = 1; row <= rowsToWrite; row++)
                            {
                                for (var column = 1; column <= colsToWrite; column++)
                                {
                                    data[row - 1, column - 1] = r.Next(99, 99999);
                                }
                            }

                            var startCell = (Range)workSheet.Cells[1, 1];
                            var endCell = (Range)workSheet.Cells[rowsToWrite, colsToWrite];
                            var writeRange = workSheet.Range[startCell, endCell];

                            writeRange.Value2 = data;

                            // second delay (if any)
                            delayFunc(timeToKeepOpen2);

                            // saving & closing file
                            workBook.SaveAs(filePath);
                            workBook.Close(true, filePath, excelAppMissing);
                            workSheet = null;
                            workBook  = null;
                            if (workSheet != null) Marshal.ReleaseComObject(workSheet);
                        }
                        
                        // ~~~~~~~~~~~~~~~~~~~ Terminating Excel instance ~~~~~~~~~~~~~~~~~~~
                        if (workBook != null) Marshal.ReleaseComObject(workBook);
                        excelApp.Quit();
                        excelApp = null;
                        if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error while opening excel file");
                        Console.WriteLine(ex.Message);
                    }
                    break;
                default:
                    Console.WriteLine("Error! Wrong file format!");
                    Console.WriteLine("Exiting...");
                    Environment.Exit(1);
                    break;
            }
        }
        // ---------------------------------------------------------------------------------------

        // -------------------------------------- Add delay --------------------------------------
        private void delayFunc(int delaySec)
        {
            if (delaySec > 0)
            {
                Console.WriteLine();
                for (int i = 0; i < delaySec; i++)
                {
                    Console.Write(" . ");
                    Thread.Sleep(1000);
                }
                Console.WriteLine();
            }
        }
        // ---------------------------------------------------------------------------------------
    }
}
