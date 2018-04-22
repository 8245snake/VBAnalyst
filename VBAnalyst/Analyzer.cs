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

        public Analyzer()
        {
            //プロシージャの定義式
            regProcedure = new Regex(@"((?<scope>(Private|Public)) )?(Sub|Function) (?<name>.+?)\(", RegexOptions.Compiled);
            //変数の定義式
            regVariable = new Regex(@"(?<scope>(Private|Public|Dim|Global))( Const)? (?!Sub|Function)(?<name>.*)", RegexOptions.Compiled);
            //代入式（If文の条件部分も拾ってしまう）
            regAssignment = new Regex(@"(?<left>.*) = (?<right>.*)", RegexOptions.Compiled);
            //End句
            regEndProc = new Regex(@"End (Sub|Function)", RegexOptions.Compiled);

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

            //結果表示
            //foreach (string item in list)
            //{
            //    Console.WriteLine(item);
            //}

            return list.ToList<string>();

        }



        //関数内を解析する（）
        public void AnalyzeProcedure(List<string> list, SourceFile sourceFile)
        {
            foreach (string item in list)
            {

            }
        }


            /// <summary>
            /// モジュール変数と関数の定義式を解析する
            /// </summary>
            /// <param name="list"></param>
            /// <returns></returns>
            private IEnumerable<string> AnalyzeModuleDefinition(IEnumerable<string> list, SourceFile sourceFile)
        {
            Match match;
            string stScope;
            string stName;
            bool blProcedureFlag = false; // 解析フラグ

            foreach (string item in list)
            {
                if (blProcedureFlag) //関数内ではEnd Subなどをチェック 
                {
                    match = regEndProc.Match(item);
                    blProcedureFlag = !match.Success;
                    yield return item;
                    continue;
                }
                else  //関数外では定義式かをチェック
                {
                    //関数の定義式かを判定
                    match = regProcedure.Match(item);
                    if (match.Success)
                    {
                        stScope = match.Groups["scope"].Value;
                        stName = match.Groups["name"].Value;
                        AnalysisData.AddProcedures(sourceFile.ID, stName, GetProcedureScope(stScope));
                        blProcedureFlag = true;
                        yield return item;
                        continue;
                    }
                }

                //関数の外での変数の定義式かを判定
                match = regVariable.Match(item);
                if (match.Success)
                {
                    stScope = match.Groups["scope"].Value;
                    stName = ExtractVariableName(match.Groups["name"].Value);
                    AnalysisData.AddVariable(sourceFile.ID, stName, GetVariableScope(stScope), KIND_MODULE);
                    continue;
                }
                else
                {
                    yield return item;
                }

            }
        }

        /// <summary>
        /// プロシージャの定義式か否かを判定
        /// </summary>
        /// <param name="stLine"></param>
        /// <returns></returns>
        private bool IsProcedureDefinition(string stLine)
        {
            return true;
        }

        /// <summary>
        /// 変数の定義式か否かを判定
        /// </summary>
        /// <param name="stLine"></param>
        /// <returns></returns>
        private bool IsVariableDefinition(string stLine)
        {
            return true;
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
    }
}
