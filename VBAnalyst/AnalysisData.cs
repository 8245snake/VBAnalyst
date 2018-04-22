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

        public static void AddVariable(Variable variable)
        {
            Variables.Add(variable);
        }

        public static void AddVariable(string stModID, string stName, int nScope, int nkind)
        {
            Variables.Add(new Variable(stModID, stName, nScope, nkind));
        }

        public static void AddProcedures(Procedure procedure)
        {
            Procedures.Add(procedure);
        }

        public static void AddProcedures(string stModID, string stName, int nScope)
        {
            Procedures.Add(new Procedure(stModID, stName, nScope));
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

    }



}
