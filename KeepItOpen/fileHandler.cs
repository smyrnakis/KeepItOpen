using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KeepItOpen
{
    public class fileHandler
    {
        // ------------------------------------ Declarations ------------------------------------
        public string fileMode = "";
        public string filePath = "";
        public int fileFormat = -1;        // 0: word , 1: excel , -1: other
        public int timeToKeepOpen1 = 1;
        public bool writeContent = false;
        public int timeToKeepOpen2 = 0;
        public bool silentMode = false;

        public bool fileIsOpen = false;

        public string[] lorems = { loremsClass.lorem1, loremsClass.lorem5, loremsClass.lorem10, loremsClass.lorem15, loremsClass.lorem20 };
        // ---------------------------------------------------------------------------------------

        public void openFile()
        {
            Console.WriteLine();
            Console.WriteLine("Entered openFile()");
            Console.WriteLine();
        }

    }
}
