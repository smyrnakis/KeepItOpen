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
        }
    }
}
