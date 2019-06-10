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

            var report = new Report();
            report.Parse(paragraphs);

            Console.WriteLine();
            Console.WriteLine(">>>> TextReport <<<<");
            Console.WriteLine();

            var docxOutputFile = @"C:\Users\dhfra\Documents\GitHub\DocConverter\TestFiles\output file.docx";

            var reportWriter = new ReportWriter(new[] { report });
            reportWriter.CreateDocX(docxOutputFile);

            //Console.WriteLine(report.TextReport());
        }
    }
}

