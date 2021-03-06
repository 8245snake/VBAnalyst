﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections;

namespace VBAnalyst
{
    /// <summary>
    /// 共通の関数、定数など
    /// </summary>
    public static class Common
    {
        public const int SCOPE_PUBLIC = 0;
        public const int SCOPE_PRIVATE = 1;

        public const int KIND_LOCAL = 0;
        public const int KIND_MODULE = 1;

        public const string FILE_FORM = "Form";
        public const string FILE_MODULE = "Module";
        public const string FILE_CLASS = "Class";

        public static int GetVariableScope(string stScope)
        {
            switch (stScope.Trim())
            {
                case "Public":
                    return SCOPE_PUBLIC;
                case "Global":
                    return SCOPE_PUBLIC;
                case "Private":
                    return SCOPE_PRIVATE;
                case "Dim":
                    return SCOPE_PRIVATE;
                default:
                    return -1;
            }
        }

        public static int GetProcedureScope(string stScope)
        {
            switch (stScope.Trim())
            {
                case "Public":
                    return SCOPE_PUBLIC;
                case "Private":
                    return SCOPE_PRIVATE;
                case "":
                    return SCOPE_PUBLIC;
                default:
                    return -1;
            }
        }

        public static string GetScopeName(int nScope)
        {
            switch (nScope)
            {
                case SCOPE_PUBLIC:
                    return "Public";
                case SCOPE_PRIVATE:
                    return "Private";
                default:
                    return "不明";
            }
        }

        public static string GetKindName(int nKind)
        {
            switch (nKind)
            {
                case KIND_LOCAL:
                    return "ローカル";
                case KIND_MODULE:
                    return "モジュール";
                default:
                    return "不明";
            }
        }

        /// <summary>
        /// ファイルを読んで1行ずつ返す
        /// </summary>
        /// <param name="stFilePath">ファイルのパス</param>
        /// <returns></returns>
        public static IEnumerable<string> ReadFileByLine(string stFilePath)
        {
            string stLine;

            using (StreamReader sr = new StreamReader(stFilePath, Encoding.GetEncoding("Shift_JIS")))
            {
                try
                {
                    while ((stLine = sr.ReadLine()) != null)
                    {
                        yield return stLine;
                    }
                }
                finally
                {
                }
            }

        }

        /// <summary>
        /// テキストを書き込む
        /// </summary>
        /// <param name="stPath">ファイルのパス</param>
        /// <param name="stText">書き込む文字列</param>
        /// <param name="blAppend">上書きか否か</param>
        public static void WriteText(string stPath, string stText, bool blAppend = false)
        {
            using (System.IO.StreamWriter sw =
                new System.IO.StreamWriter(stPath, blAppend, System.Text.Encoding.GetEncoding("shift_jis")))
            {
                try
                {
                    sw.WriteLine(stText);
                }
                catch
                {
                }
            }
        }





    }


}
