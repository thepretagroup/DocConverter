using System;
namespace DocConverter
{
    class Sample
    {
        [STAThread]
        static void Main(string[] args)
        {
            var rtfFile = @"C:\Users\dhfra\Documents\GitHub\DocConverter\TestFiles\input file.rtf";

            var paragraphs = RtfReader.ReadParagraphs(rtfFile);

            var report = new Report(paragraphs);

            Console.WriteLine();
            Console.WriteLine(">>>> TextReport <<<<");
            Console.WriteLine();

            Console.WriteLine(report.TextReport());
        }
    }
}

