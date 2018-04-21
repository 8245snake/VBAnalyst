using static System.IO.Path;
using static VBAnalyst.Common;

namespace VBAnalyst
{
    /// <summary>
    /// SourceFileオブジェクトのファクトリークラス
    /// </summary>
    public static class SourceFileFactory
    {
        private static long ClIdMax = 0;

        /// <summary>
        /// VBPの1行を渡してオブジェクトを作る。作成できなければnullを返す。
        /// </summary>
        /// <param name="stLine"></param>
        /// <returns></returns>
        public static SourceFile CreateSourceFile(string stLine, string stDir)
        {
            string[] stToken;

            stToken = stLine.Split('=');
            if (stToken.Length < 2) { return null; }

            switch (stToken[0].Trim())
            {
                case FILE_FORM:
                    return new FrmFile(ClIdMax++, GetAbsolutePath(stDir, stToken[1].Trim()));
                case FILE_MODULE:
                    string[] stNames;
                    stNames = stToken[1].Split(';');
                    return new BasFile(ClIdMax++, stNames[0].Trim(), GetAbsolutePath(stDir, stNames[1].Trim()));
                case FILE_CLASS:
                    return new ClsFile(ClIdMax++, GetAbsolutePath(stDir, stToken[1].Trim()));
                default:
                    return null;
            }

        }

        /// <summary>
        /// パスを連結して絶対パスに変換
        /// </summary>
        /// <param name="stDir">ディレクトリ</param>
        /// <param name="stPath">相対パス</param>
        /// <returns>失敗したら空文字を返す</returns>
        private static string GetAbsolutePath(string stDir, string stPath)
        {
            try
            {
                return GetFullPath(Combine(stDir, stPath));
            }
            catch
            {
                return "";
            }
        }

    }
}
