using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VBAnalyst
{
    public class BasFile : SourceFile
    {

        public BasFile(long lID, string stName, string stFilePath)
        {
            CstModuleID = CreateModuleID(lID);
            CstModuleName = stName;
            CstFilePath = stFilePath;
        }


    }
}
