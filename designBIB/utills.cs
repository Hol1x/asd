using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace designBIB
{
    class utills
    {
        public static bool FileCheck(string fileToCheck) {
            return File.Exists(fileToCheck);
        }
    }
}
