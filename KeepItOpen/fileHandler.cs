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
        public string openMode = "";
        public string filePath = "";
        public int fileFormat = -1;        // 0: word , 1: excel , -1: other
        public int timeToKeepOpen1 = 0;
        public int timeToKeepOpen2 = 0;
        public bool silentMode = false;
        
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

                            wordFile.SaveAs2(filePath);
                            wordFile.Close(false, ref wordAppMissing, ref wordAppMissing);
                            wordFile = null;
                        }

                        // ~~~~~~~~~~~~~~~~~~~~ Terminating Word instance ~~~~~~~~~~~~~~~~~~~
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
                        excelApp.Visible = false;
                        excelApp.DisplayAlerts = false;
                        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine("   ~~~ Excel part is not implemented yet! ~~~   ");
                        Console.WriteLine("   ~~~ Please wait for the next release! ~~~   ");
                        Console.WriteLine();
                        Console.WriteLine();

                        // ~~~~~~~~~~~~~~~~~~~ Terminating Excel instance ~~~~~~~~~~~~~~~~~~~
                        excelApp.Quit();
                        if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
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
