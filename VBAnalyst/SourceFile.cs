using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VBAnalyst
{
    public abstract class SourceFile
    {
        protected string CstModuleID; //一意のモジュールID
        protected string CstModuleName; //モジュール名
        protected string CstFilePath; //絶対パス
        protected List<string> DefinedVariables; //このモジュール内で定義している変数リスト
        protected List<string> DefinedProcedures; //このモジュール内で定義しているプロシージャリスト

        protected List<string> UsingVariables; //このモジュール内で使用している変数リスト
        protected List<string> UsingProcedures; //このモジュール内で使用しているプロシージャリスト

        public string ID { get { return CstModuleID; } }
        public string Name { get { return CstModuleName; } }
        public string Path { get { return CstFilePath; } }

        protected string CreateModuleID(long lID)
        {
            return "MOD" + string.Format("{0:0000}",lID);
        }

        public void AddDefinedVariable(string stID)
        {
            DefinedVariables.Add(stID);
        }

        public void AddDefinedProcedure(string stID)
        {
            DefinedProcedures.Add(stID);
        }

        public override string ToString()
        {
            return string.Format("ID:{0}, Name:{1}", CstModuleID, CstModuleName);
        }


    }
}
