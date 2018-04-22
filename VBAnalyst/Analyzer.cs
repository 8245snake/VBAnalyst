using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static VBAnalyst.Common;
using System.Text.RegularExpressions;

namespace VBAnalyst
{
    public class Analyzer
    {
        readonly string[] IgnoreList = { "Attribute", "Option", "VERSION" };
        Regex regProcedure; //プロシージャ
        Regex regVariable; //変数定義
        Regex regAssignment; //代入式
        Regex regEndProc; //End SubもしくはEnd Function
        Regex regTypeDef; //ユーザー定義型
        Regex regEnum; //列挙体

        public Analyzer()
        {
            //プロシージャの定義式
            regProcedure = new Regex(@"((?<scope>(Private|Public)) )?(Sub|Function) (?<name>.+?)\(", RegexOptions.Compiled);
            //変数の定義式
            regVariable = new Regex(@"(?<scope>(Private|Public|Dim|Global|Const))( Const)? (?!Sub|Function|Type|Enum)(?<name>.*)", RegexOptions.Compiled);
            //代入式（If文の条件部分も拾ってしまう）
            regAssignment = new Regex(@"(?<left>.*) = (?<right>.*)", RegexOptions.Compiled);
            //End句
            regEndProc = new Regex(@"End (Sub|Function)", RegexOptions.Compiled);
            //ユーザー定義型
            regTypeDef = new Regex(@"(?<scope>Private|Public)? ?Type (?<name>.*)", RegexOptions.Compiled);
            //列挙体
            regEnum = new Regex(@"(?<scope>Private|Public)? ?Enum (?<name>.*)", RegexOptions.Compiled);
        }

        /// <summary>
        /// 関数外で定義している変数と関数を登録したら一旦リストを返す
        /// </summary>
        /// <param name="sourceFile"></param>
        public List<string> AnalyzeModule(SourceFile sourceFile)
        {
            IEnumerable<string> list;

            //余計な要素を削除
            list = DeleteText(ReadFileByLine(sourceFile.Path));
            list = ConmbineUnderscore(list);
            list = DeleteBiginEnd(list);
            list = ApplyIgnoreList(list);

            //関数外での定義を解析
            list = AnalyzeModuleDefinition(list, sourceFile);

            ////結果表示
            //foreach (string item in list)
            //{
            //    Console.WriteLine(item);
            //}

            return list.ToList<string>();

        }

        //関数内で参照している変数、関数を解析する
        public void AnalyzeProcedure(List<string> list, SourceFile sourceFile)
        {
            Match match;
            string stName;

            foreach (string item in list)
            {
                //End Sub or Functionなら戻る
                match = regEndProc.Match(item);
                if (match.Success)
                {
                    continue;
                }

                //関数の定義式かを判定（stName更新）
                match = regProcedure.Match(item);
                if (match.Success)
                {
                    stName = match.Groups["name"].Value;
                    //名前とモジュールIDをキーにしてプロシージャIDを取得したい
                }



            }
        }


        /// <summary>
        /// モジュール変数と関数の定義式を解析する。関数内のローカル変数も取得する？
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private IEnumerable<string> AnalyzeModuleDefinition(IEnumerable<string> list, SourceFile sourceFile)
        {
            Match match;
            string stScope;
            string stName;
            bool blProcedureFlag = false; // 関数解析フラグ
            bool blTypeFlag = false; // ユーザー定義型解析フラグ
            bool blEnumFlag = false; // 列挙体解析フラグ
            string stEnumScope = "";
            string stProcedureID = "";

            foreach (string item in list)
            {
                if (blProcedureFlag) //関数内の解析
                {
                    //ここでローカル変数も取得したい




                    //Endかを判定
                    match = regEndProc.Match(item);
                    blProcedureFlag = !match.Success;
                    yield return item;
                    continue;
                }

                if (blEnumFlag) //Enum内の解析 
                {
                    if (item == "End Enum") //Endかを判定
                    {
                        blEnumFlag = false;
                    }
                    else
                    {
                        //Enumの要素1つ1つを変数に登録
                        stName = ExtractVariableName(item);
                        AnalysisData.AddVariable(sourceFile.ID, stName, GetVariableScope(stEnumScope), KIND_MODULE);
                    }
                    continue;
                }

                if (blTypeFlag) //ユーザー定義型内の解析
                {
                    if (item == "End Type") //Endかを判定
                    {
                        blTypeFlag = false;
                    }
                    continue;
                }

                //関数の定義式かを判定
                match = regProcedure.Match(item);
                if (match.Success)
                {
                    stScope = match.Groups["scope"].Value;
                    stName = match.Groups["name"].Value;
                    stProcedureID = AnalysisData.AddProcedure(sourceFile.ID, stName, GetProcedureScope(stScope));

                    //ここで引数も取得
                    RegArgVariable(stProcedureID,item);

                    blProcedureFlag = true;
                    yield return item;
                    continue;
                }

                //変数の定義式かを判定
                match = regVariable.Match(item);
                if (match.Success)
                {
                    stScope = match.Groups["scope"].Value;
                    stName = ExtractVariableName(match.Groups["name"].Value);
                    AnalysisData.AddVariable(sourceFile.ID, stName, GetVariableScope(stScope), KIND_MODULE);
                    continue;
                }

                //列挙体の定義式かを判定
                match = regEnum.Match(item);
                if (match.Success)
                {
                    stEnumScope = match.Groups["scope"].Value;
                    stName = ExtractVariableName(match.Groups["name"].Value);
                    AnalysisData.AddVariable(sourceFile.ID, stName, GetVariableScope(stEnumScope), KIND_MODULE);
                    blEnumFlag = true;
                    continue;
                }

                //ユーザー定義型の定義式かをチェック
                match = regTypeDef.Match(item);
                if (match.Success)
                {
                    stScope = match.Groups["scope"].Value;
                    stName = match.Groups["name"].Value;
                    AnalysisData.AddType(sourceFile.ID, stName, GetProcedureScope(stScope));
                    blTypeFlag = true;
                    continue;
                }

                yield return item;
            }
        }

        /// <summary>
        /// コメントや文字列リテラルを消す
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private IEnumerable<string> DeleteText(IEnumerable<string> list)
        {
            string stCleanLine;

            foreach (string stLine in list)
            {
                stCleanLine = CleanUpLine(stLine);

                if (stCleanLine != "")
                {
                    yield return stCleanLine;
                }
                else
                {
                    continue;
                }
            }
        }

        /// <summary>
        /// アンダースコアでつながった行を結合して返す
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private IEnumerable<string> ConmbineUnderscore(IEnumerable<string> list)
        {
            string stCleanLine;
            string stLineBuff = "";
            bool blMultiFlag = false; // 解析フラグ

            foreach (string stLine in list)
            {
                stCleanLine = stLine;

                //末端が「_」か否か
                if (stLine.Length > 0 && stLine.Last() == '_')
                {
                    blMultiFlag = true;
                }

                if (blMultiFlag)
                {
                    if (stLine.Length > 0 && stLine.Last() == '_')
                    {
                        stLineBuff += stLine.Trim('_');
                        continue;
                    }
                    else
                    {
                        stLineBuff += stLine.Trim();
                        stCleanLine = stLineBuff;
                        stLineBuff = "";
                        blMultiFlag = false;
                    }
                }

                yield return stCleanLine;
            }
        }

        /// <summary>
        /// Bigin-Endを削除して返す
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private IEnumerable<string> DeleteBiginEnd(IEnumerable<string> list)
        {
            int nBiginCount = 0;   //Biginで増えてEndで減る（frm解析用）

            foreach (string stLine in list)
            {
                if (stLine.IndexOf("Begin") == 0)
                {
                    nBiginCount++;
                    continue;
                }

                if (nBiginCount > 0)
                {
                    if (stLine.IndexOf("End") == 0)
                    {
                        nBiginCount--;
                    }
                    continue;
                }

                yield return stLine;
            }
        }

        /// <summary>
        /// 無視リストに含まれる文字列から始まる行を除く
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private IEnumerable<string> ApplyIgnoreList(IEnumerable<string> list)
        {
            bool blFind = false;

            foreach (string stLine in list)
            {
                blFind = false;
                for (int i = 0; i < IgnoreList.Length; i++)
                {
                    if (stLine.IndexOf(IgnoreList[i]) == 0)
                    {
                        blFind = true;
                        break;
                    }
                }
                if (!blFind) { yield return stLine; }
            }
        }

        /// <summary>
        /// コメントと文字列を削除し、連続した半角スペースを1つにする
        /// </summary>
        /// <param name="stLine">ソースコードの1行</param>
        /// <returns></returns>
        private string CleanUpLine(string stLine)
        {
            char[] stArr = stLine.Trim().ToCharArray();
            bool blIsString = false;

            string stRtn = "";

            for (int i = 0; i < stArr.Length; i++)
            {
                //文字列の中で現れる「'」は無視する
                if (stArr[i] == '"')
                {
                    blIsString = !blIsString;
                    stRtn += stArr[i];
                    continue;
                }
                if (blIsString) { continue; }

                //半角スペースが2回以上続かないようにする
                if (stArr[i] == ' ' && i > 0)
                {
                    if (stArr[i - 1] == ' ') { continue; }
                }

                //「'」が出たら終了
                if (stArr[i] == '\'') { return stRtn; }
                stRtn += stArr[i];
            }

            return stRtn;
        }

        /// <summary>
        /// 文字列の半角スペースか左括弧までを取得
        /// </summary>
        /// <param name="stName"></param>
        /// <returns></returns>
        private string ExtractVariableName(string stName)
        {
            char[] stArr = stName.Trim().ToCharArray();
            string stRtn = "";

            foreach (char item in stArr)
            {
                if (item == ' ' || item == '(')
                {
                    return stRtn;
                }

                stRtn += item;
            }
            return stRtn;
        }

        /// <summary>
        /// プロシージャ定義式から引数を抜き出して登録する（最初のカッコから最後のカッコまでの文字列を取得すればOK?）
        /// </summary>
        /// <param name="stProcedureID"></param>
        /// <param name="stLine"></param>
        private void RegArgVariable(string stProcedureID, string stLine)
        {
            //エラーチェック
            if (stProcedureID == "" || stLine == "") { return; }

            int nFirstIndex = stLine.IndexOf('(');
            int nLastIndex = stLine.LastIndexOf(')');

            if (nFirstIndex < 0 || nLastIndex < 0) { return; }

            string stDefinition = stLine.Substring(nFirstIndex + 1, nLastIndex - nFirstIndex - 1).Trim();

            if (stDefinition == "") { return; }

            string[] stDefArr = stDefinition.Split(',');
            string[] stVar;
            string stVariable; 
            foreach (string item in stDefArr)
            {
                stVariable = item;
                stVariable = stVariable.Replace("ByVal", "");
                stVariable = stVariable.Replace("ByRef", "");
                stVariable = stVariable.Replace("Optional", "");
                stVariable = stVariable.Replace("()", "");
                stVar = stVariable.Split(' ');
                stVariable = stVar[0].Trim();
                AnalysisData.AddVariable(stProcedureID, stVariable, SCOPE_PRIVATE, KIND_LOCAL);
            }
        }
    }
}
