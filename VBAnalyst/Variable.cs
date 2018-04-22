using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static VBAnalyst.Common;

namespace VBAnalyst
{
    /// <summary>
    /// 変数クラス
    /// </summary>
    public class Variable
    {
        private string CstVariableID; //変数ID
        private string CstModuleID; //この変数を定義しているモジュールのID（ローカルだったらプロシージャID入れる？）
        private string CstName; ///変数名
        private int CnScope; //スコープ（SCOPE_PUBLIC or SCOPE_PRIVATE）
        private int CnKind; //種類（モジュール変数 or ローカル変数）

        private static long ClIdMax = 0;

        public string ID { get { return CstVariableID; } }
        public string Name { get { return CstName; } }

        public Variable(string stModID,string stName,int nScope, int nkind)
        {
            CstVariableID = CreateVariableID();
            CstModuleID = stModID;
            CstName = stName;
            CnScope = nScope;
            CnKind = nkind;
        }

        private string CreateVariableID()
        {
            return "VAR" + string.Format("{0:00000000}", ClIdMax++);
        }

        public override string ToString()
        {
            return string.Format("ID:{0}, 名前:{1}, スコープ:{2}, 種類:{3}", CstVariableID, CstName, GetScopeName(CnScope),GetKindName(CnKind));
        }
    }
}
