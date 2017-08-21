/**************************************************************************************
    KeepItOpen
    Copyright (C) 2017  Apostolos Smyrnakis - IT/CDA/AD - apostolos.smyrnakis@cern.ch

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 **************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Text.RegularExpressions;
using System.IO;
using KeepItOpen;

namespace keepItOpen
{
    public class Program
    {
        // ------------------------------------ MAIN function ------------------------------------
        [STAThread]
        static void Main(string[] args)
        {
            displayArguments dArgs = new displayArguments();
            fileHandler fHand = dArgs.getFileHandler();

            if (args.Length == 0)
            {
                Console.WriteLine("Please give arguments!");
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Available arguments ~~~~~~~~~~~~~~~~~~~~~~~~");
                Console.WriteLine();
                Console.WriteLine(" 'r' or 'w' : file in 'read' or 'write' mode");
                Console.WriteLine(" file + <path> : path to word OR excel file to keep open");
                Console.WriteLine(" delay1 + <integer[1-3600]> : delay in seconds that file will be kept open");
                Console.WriteLine(" delay2 + <integer[0-3600]> : extra delay in seconds (ONLY in 'w' mode)");
                Console.WriteLine(" 's' : silent mode - no confirmation asked");
                Console.WriteLine();
                Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("Press any key to exit...");
                Console.Read();
                Environment.Exit(1);
            }
            else     // ------------------------------- Get arguments and check validity -------------------------------
            {
                try
                {
                    for (int argIdx = 0; argIdx < args.Length; argIdx++)
                    {
                        switch (args[argIdx])
                        {
                            case "r":
                                fHand.openMode = "r";
                                break;
                            case "w":
                                fHand.openMode = "w";
                                break;
                            case "file":
                                fHand.filePath = Path.GetFullPath(@args[argIdx + 1]);
                                break;
                            case "delay1":
                                fHand.timeToKeepOpen1 = Convert.ToInt32(args[argIdx + 1]);
                                break;
                            case "delay2":
                                fHand.timeToKeepOpen2 = Convert.ToInt32(args[argIdx + 1]);
                                break;
                            case "s":
                                fHand.silentMode = true;
                                break;
                            default:
                                break;
                        }
                    }
                    
                    // ~~~~~~~~~~~~~~~~~~~~ Checking data validity ~~~~~~~~~~~~~~~~~~~~
                    // ~~ r/w mode
                    if ((fHand.openMode != "r") && (fHand.openMode != "w"))
                    {
                        Console.WriteLine();
                        Console.WriteLine("Error in 'mode' argument!");
                        Console.WriteLine("'r' or 'w' : file in 'read' or 'write' mode");
                        Console.WriteLine();
                        Console.Read();
                        Environment.Exit(1);
                    }
                    // ~~ file path (and file type)
                    if (File.Exists(fHand.filePath))
                    {
                        string fileExtension = Path.GetExtension(fHand.filePath);             // 0: word , 1: excel , -1: other
                        if ((fileExtension == ".doc") || (fileExtension == ".docx"))
                            fHand.fileFormat = 0;   //  0 : word
                        else if ((fileExtension == ".xls") || (fileExtension == ".xlsx"))
                            fHand.fileFormat = 1;   //  1 : excel
                        else
                        {
                            fHand.fileFormat = -1;  // -1 : other
                            Console.WriteLine();
                            Console.WriteLine("Wrong file format! Exiting application...");
                            Console.WriteLine();
                            Console.Read();
                            Environment.Exit(1);
                        }
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("Error! Path does not exist!");
                        Console.WriteLine("Selected path: " + fHand.filePath.ToString());
                        Console.WriteLine("file + <path> : path to word OR excel file to keep open");
                        Console.WriteLine();
                        Console.Read();
                        Environment.Exit(1);
                    }
                    // ~~ Getting 1st time delay in seconds
                    if ((fHand.timeToKeepOpen1 < 1) || (fHand.timeToKeepOpen1 > 3600))
                    {
                        Console.WriteLine();
                        Console.WriteLine("Error in 1st delay argument!");
                        Console.WriteLine("delay1 + <integer[1 - 3600]> : delay in seconds that file will be kept open");
                        Console.WriteLine();
                        Console.Read();
                        Environment.Exit(1);
                    }
                    // ~~ Getting 2nd time delay in seconds
                    if ((fHand.timeToKeepOpen2 < 0) || (fHand.timeToKeepOpen2 > 3600))
                    {
                        Console.WriteLine();
                        Console.WriteLine("Error in 2nd delay argument!");
                        Console.WriteLine("delay2 + <integer[0 - 3600]> : extra delay in seconds that file will be kept open");
                        Console.WriteLine();
                        Console.Read();
                        Environment.Exit(1);
                    }
                    else
                    {
                        if ((fHand.openMode == "r") && (fHand.timeToKeepOpen2 != 0))
                        {
                            Console.WriteLine();
                            Console.WriteLine("READ mode selected. Extra delay ('delay2') will be ignored!");
                        }
                    }
                    // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error reading arguments! Check the order!");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine();
                    Console.Read();
                    Environment.Exit(1);
                }
            }       // ------------------------------------------------------------------------------------------------

            // If NOT 'silent mode' display arguments, else continue with the program
            if (!fHand.silentMode)
                dArgs.displayArgs();
            else
                fHand.openFile();

            Console.WriteLine("Run once more? (y/n)");
            string continueYN = Console.ReadLine();
            if (continueYN.ToLower() == "n")
            {
                Console.WriteLine();
                Console.WriteLine("Exiting ...");
                Environment.Exit(0);
            }
            else if (continueYN.ToLower() == "y")
            {
                fHand.openFile();
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("Wrong input!");
            }

            Console.WriteLine();
            Console.WriteLine("Press 'enter' to exit...");
            Console.Read();
            Environment.Exit(0);        // Exiting
        }
        // ---------------------------------------------------------------------------------------
    }
}
