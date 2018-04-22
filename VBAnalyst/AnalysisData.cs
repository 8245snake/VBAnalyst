using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace VBAnalyst
{
    /// <summary>
    /// データクラス
    /// </summary>
    public class AnalysisData
    {
        public static List<Variable> Variables = new List<Variable>(); //変数リスト
        public static List<Procedure> Procedures = new List<Procedure>(); //プロシージャリスト
        public static List<UserDefType> Types = new List<UserDefType>(); //ユーザー定義型リスト

        //変数
        public static void AddVariable(Variable variable)
        {
            Variables.Add(variable);
        }

        public static string AddVariable(string stModID, string stName, int nScope, int nkind)
        {
            Variable variable = new Variable(stModID, stName, nScope, nkind);
            Variables.Add(variable);
            return variable.ID;
        }

        //プロシージャ
        public static void AddProcedure(Procedure procedure)
        {
            Procedures.Add(procedure);
        }

        public static string AddProcedure(string stModID, string stName, int nScope)
        {
            Procedure proc = new Procedure(stModID, stName, nScope);
            Procedures.Add(proc);
            return proc.ID;
        }

        //ユーザー定義型
        public static void AddType(UserDefType userDefType)
        {
            Types.Add(userDefType);
        }

        public static void AddType(string stModID, string stName, int nScope)
        {
            Types.Add(new UserDefType(stModID, stName, nScope));
        }


        public static void EnumVariables()
        {
            foreach (Variable item in Variables)
            {
                Console.WriteLine(item.ToString());
            }
        }

        public static void EnumProcedures()
        {
            foreach (Procedure item in Procedures)
            {
                Console.WriteLine(item.ToString());
            }
        }

        public static void EnumTypes()
        {
            foreach (UserDefType item in Types)
            {
                Console.WriteLine(item.ToString());
            }
        }

    }



}
