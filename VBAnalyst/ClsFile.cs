using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static VBAnalyst.Common;

namespace VBAnalyst
{
    public class ClsFile : SourceFile
    {
        public ClsFile(long lID, string stFilePath)
        {
            CstModuleID = CreateModuleID(lID);
            CstFilePath = stFilePath;


        }


    }
}
