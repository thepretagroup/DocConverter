using System;
using System.Linq;
using System.Text;
using GemBox.Document;

namespace DocConverter
{
    class Sample
    {
        [STAThread]
        static void Main(string[] args)
        {
            var rtfFile = @"C:\Users\dhfra\Documents\GitHub\DocConverter\input file.rtf";

            var paragraphs = RtfReader.ReadParagraphs(rtfFile);

            var report = new Report(paragraphs);

            Console.WriteLine(report.ToString());

            //foreach (var line in paragraphs)
            //{
            //    Console.WriteLine(line);
            //}

            //RTFtoTxt();
        }

    }
}

