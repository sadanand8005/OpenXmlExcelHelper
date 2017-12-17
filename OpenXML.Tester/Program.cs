using OpenXML.Tester.Test;
using System;
using System.Diagnostics;

namespace OpenXML.Tester
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //new ZipTester().Test();

                TimeAndMeasureGC("Excel Tester. Initial Test", () =>
                {
                    new ExcelXmlHelper().Test().GetAwaiter().GetResult();
                });

                TimeAndMeasureGC("Excel Tester. Test 1", () =>
                {
                    new ExcelXmlHelper().Test().GetAwaiter().GetResult();
                });


                TimeAndMeasureGC("Excel Tester. Test 2", () =>
                {
                    new ExcelXmlHelper().Test().GetAwaiter().GetResult();
                });

                TimeAndMeasureGC("Excel Tester. Test 3", () =>
                {
                    new ExcelXmlHelper().Test().GetAwaiter().GetResult();
                });


                Console.ReadKey();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }

        static int[] GetGCPerGenCounts()
        {
            int[] counts = new int[3];
            for (int i = 0; i < 3; ++i)
            {
                counts[i] = GC.CollectionCount(i);
            }
            return counts;
        }

        static void TimeAndMeasureGC(string what, Action action)
        {
            int[] gcCountsBefore = GetGCPerGenCounts();
            Stopwatch sw = Stopwatch.StartNew();
            action();
            int[] gcCountsAfter = GetGCPerGenCounts();
            Console.Write("{0}\tTime = {1}ms\t", what, sw.ElapsedMilliseconds);
            for (int i = 0; i < gcCountsAfter.Length; ++i)
            {
                Console.Write("Gen{0} GCs = {1}\t", i, gcCountsAfter[i] - gcCountsBefore[i]);
            }
            Console.Write("Mem = {0:0.00}MB", GC.GetTotalMemory(true) / 1024576.0);
            Console.WriteLine();
        }


    }
}
