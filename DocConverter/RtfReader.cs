﻿using System.Collections.Generic;

namespace DocConverter
{
    public static class RtfReader
    {
        public static IList<string> ReadParagraphs(string rtfFile)
        {
            return RichTextBoxReadParagraphs(rtfFile);
        }

        public static IList<string> RichTextBoxReadParagraphs(string rtfFile)
        {
            // Based on example in https://social.msdn.microsoft.com/Forums/vstudio/en-US/6e56af9b-d7d3-49f3-9ec4-80edde3fe54b/reading-modifying-rtf-files?forum=csharpgeneral
            System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();
            string rtfText = System.IO.File.ReadAllText(rtfFile);
            rtBox.Rtf = rtfText;
            string plainText = rtBox.Text;

            //Console.WriteLine("%%%%  RichTextBoxReadParagraphs");
            //Console.WriteLine(plainText);

            var paragraphs = plainText.Split('\n');
            return paragraphs;
        }
    }
}
