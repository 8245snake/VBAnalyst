using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static VBAnalyst.Common;

namespace VBAnalyst
{
    public class Procedure
    {
        private string CstProcedureID; //プロシージャID
        private string CstModuleID; //このプロシージャを定義しているモジュールのID
        private string CstName; ///プロシージャ名
        private int CnScope; //スコープ（SCOPE_PUBLIC or SCOPE_PRIVATE）

        private static long ClIdMax = 0;

        public Procedure(long lID, string stModID, string stName, int nScope)
        {
            CstProcedureID = CreateProcedureID();
            CstModuleID = stModID;
            CstName = stName;
            CnScope = nScope;
        }

        private string CreateProcedureID()
        {
            return "PRC" + string.Format("{0:00000000}", ClIdMax++);
        }

        public override string ToString()
        {
            return string.Format("プロシージャ名:{0}, スコープ:{1}", CstName, GetScopeName(CnScope));
        }
    }


}
