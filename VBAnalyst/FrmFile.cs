using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static VBAnalyst.Common;

namespace VBAnalyst
{
    public class FrmFile : SourceFile
    {
        public FrmFile(long lID, string stFilePath)
        {
            CstModuleID = CreateModuleID(lID);
            CstFilePath = stFilePath;
            CstModuleName = GetModuleName();
        }

        private string GetModuleName()
        {
            foreach (string stLine in ReadFileByLine(CstFilePath))
            {
                if (stLine.Contains("Begin VB.Form"))
                {
                    return stLine.Replace("Begin VB.Form", "").Trim();
                }
            }
            return "";
        }
    }
}
