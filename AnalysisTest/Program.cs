using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static VBAnalyst.Common;
using VBAnalyst;

namespace AnalysisTest
{
    class Program
    {
        static void Main(string[] args)
        {

            string stDir = @"C:\Users\Shingo\Downloads\GSPL88J101\";
            SourceFile file;

            Analyzer analyzer = new Analyzer();
            List<string> list;

            var stopwatch = new System.Diagnostics.Stopwatch();


            // 計測開始
            stopwatch.Start();
            foreach (string stLine in ReadFileByLine(stDir + "GPSPL88J_IF.vbp"))
            {
                file = SourceFileFactory.CreateSourceFile(stLine, stDir);

                if (file != null)
                {
                    Console.WriteLine(file.ToString());
                    list = analyzer.AnalyzeModule(file);
                    Console.WriteLine("");
                }
            }
            // 計測停止
            stopwatch.Stop();

            //TimeSpan timespan = stopwatch.Elapsed;
            Console.WriteLine($"　{stopwatch.ElapsedMilliseconds}ミリ秒");

            AnalysisData.EnumVariables();
            AnalysisData.EnumProcedures();

            Console.ReadKey();
        }
    }
}
