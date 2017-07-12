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

namespace keepItOpen
{
    class Program
    {
        bool fileOpen = false;

        string fileMode = "";
        string filePath = "";
        int timeToKeepOpen = 1;

        private void Main(string[] args)
        {
            if (args.Length <= 0)
            {
                Console.WriteLine("Please give arguments!");
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~ Available arguments ~~~~~~~~~~~~~~~~~~~~~~~");
                Console.WriteLine();
                Console.WriteLine(" 'r' or 'w' : file in 'read' or 'write' mode");
                Console.WriteLine(" <path> : path to word OR excel file to keep open");
                Console.WriteLine(" <integer[1-3600]> : delay in seconds that file will be kept open");
                Console.WriteLine();
                Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                Console.WriteLine();
                Console.WriteLine();
                Environment.Exit(1);
            }
            else
            {
                try
                {
                    // ~~~ Getting mode (read or write)
                    if ((args[1].ToLower() == "r") || (args[1].ToLower() == "w"))
                    {
                        fileMode = args[1];
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("Error in the first argument!");
                        Console.WriteLine("'r' or 'w' : file in 'read' or 'write' mode");
                        Console.WriteLine();
                        Environment.Exit(1);
                    }

                    // ~~~ Getting file path (and file type)

                    if ()
                    {

                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("Error in the second argument!");
                        Console.WriteLine("<path> : path to word OR excel file to keep open");
                        Console.WriteLine();
                        Environment.Exit(1);
                    }

                    // ~~~ Getting time in seconds
                    if ((Convert.ToInt32(args[3]) >= 1) && (Convert.ToInt32(args[3]) <= 3600))
                    {
                        timeToKeepOpen = Convert.ToInt32(args[3]);
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("Error in the third argument!");
                        Console.WriteLine("<integer[1 - 3600]> : delay in seconds that file will be kept open");
                        Console.WriteLine();
                        Environment.Exit(1);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error reading arguments!");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine();
                    Environment.Exit(1);
                }

            }
        }
        // ---------------------------------------------------------------------------------------


    }
}
