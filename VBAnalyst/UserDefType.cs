using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static VBAnalyst.Common;

namespace VBAnalyst
{
    /// <summary>
    /// ユーザー定義型（構造体）
    /// </summary>
    public class UserDefType
    {
        private string CstTypeID; //型ID
        private string CstModuleID; //この型を定義しているモジュールのID
        private string CstName; ///型名
        private int CnScope; //スコープ（SCOPE_PUBLIC or SCOPE_PRIVATE）

        private static long ClIdMax = 0;

        public UserDefType(string stModID, string stName, int nScope)
        {
            CstTypeID = CreateVariableID();
            CstModuleID = stModID;
            CstName = stName;
            CnScope = nScope;
        }

        private string CreateVariableID()
        {
            return "TYP" + string.Format("{0:00000000}", ClIdMax++);
        }

        public override string ToString()
        {
            return string.Format("ID:{0}, 名前:{1}, スコープ:{2}", CstTypeID, CstName, GetScopeName(CnScope));
        }
    }
}
