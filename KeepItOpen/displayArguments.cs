using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace KeepItOpen
{
    public class displayArguments
    {
        private fileHandler fHand = new fileHandler();
        public fileHandler getFileHandler()
        {
            return fHand;
        }

        public void displayArgs()
        {
            // -- Temp strings for data presenting --
            string tempOpenMode = "";
            if (fHand.openMode == "r")
                tempOpenMode = "READ";
            else
                tempOpenMode = "WRITE";
            // --------------------------------------
            string tempFileFormat = "";
            if (fHand.fileFormat == 0)
                tempFileFormat = "Word file";
            else if (fHand.fileFormat == 1)
                tempFileFormat = "Excel file";
            else
                tempFileFormat = "other file";
            // --------------------------------------
            string tempWriteContent = "";
            if (fHand.openMode == "w")
                tempWriteContent = " BE";
            else
                tempWriteContent = " NOT be";
            // --------------------------------------
            string tempExtraTime = "";
            if ((fHand.timeToKeepOpen2 > 0) && (fHand.openMode == "w"))
                tempExtraTime = " Extra delay of "
                    + fHand.timeToKeepOpen2.ToString()
                    + " seconds will be added before closing.";
            // --------------------------------------

            // ~~~ TEMPORARY - EXCEL NOT READY YET! ~~~
            if (fHand.fileFormat == 1)
            {
                Console.WriteLine();
                Console.WriteLine();
                Console.WriteLine("   ~~~ Excel part is not implemented yet! ~~~   ");
                Console.WriteLine("   ~~~ Please wait for the next release! ~~~   ");
                Console.WriteLine();
                Console.WriteLine();
                return;
            }
            // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

            Console.WriteLine();
            Console.WriteLine(" Opening " + tempFileFormat + " in " + tempOpenMode + " mode.");
            Console.WriteLine(" File will be kept open for " + fHand.timeToKeepOpen1.ToString() + " seconds.");
            Console.WriteLine(" Content will" + tempWriteContent + " written to the file!");
            Console.WriteLine(tempExtraTime);
            Console.WriteLine(" File path: " + fHand.filePath);
            Console.WriteLine();
            Console.WriteLine("Continue? (y/n)");
            string continueYN = Console.ReadLine();
            if (continueYN.ToLower() == "n")
            {
                Console.WriteLine();
                Console.WriteLine("Exiting ...");
                Environment.Exit(1);
            }
            else if (continueYN.ToLower() == "y")
            {
                fHand.openFile();
            }
            else if (continueYN.ToLower() != "y")
            {
                Console.WriteLine();
                Console.WriteLine("Wrong input!");
            }
        }
    }
}
