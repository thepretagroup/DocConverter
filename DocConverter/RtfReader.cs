using GemBox.Document;
using System;
using System.Collections.Generic;
using System.Linq;


namespace DocConverter
{
    public static class RtfReader
    {
        public static IList<string> ReadParagraphs(string rtfFile)
        {
            //return GemBoxReadParagraphs(rtfFile);
            return RichTextBoxReadParagraphs(rtfFile);
        }

        public static IList<string> RichTextBoxReadParagraphs(string rtfFile)
        {
            // Based on example in https://social.msdn.microsoft.com/Forums/vstudio/en-US/6e56af9b-d7d3-49f3-9ec4-80edde3fe54b/reading-modifying-rtf-files?forum=csharpgeneral
            System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();
            string rtfText = System.IO.File.ReadAllText(rtfFile);
            rtBox.Rtf = rtfText;
            string plainText = rtBox.Text;

            Console.WriteLine("%%%%  RichTextBoxReadParagraphs");
            Console.WriteLine(plainText);

            //using (System.IO.StreamWriter file =
            //    new System.IO.StreamWriter(@"C:\test\WriteLines.rtf"))
            //{
            //    file.WriteLine(plainText);
            //}

            //var paragraphs = plainText.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            var paragraphs = plainText.Split('\n');
            return paragraphs;
        }

        public static IList<string> GemBoxReadParagraphs(string rtfFile)
        {
            var paragraphs = new List<string>();
            DocumentModel document = null;

            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // Continue to use the component in a Trial mode when free limit is reached.
            ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

            try
            {
                //document = DocumentModel.Load(@"C:\Users\dhfra\Documents\GitHub\DocConverter\input file.rtf");
                document = DocumentModel.Load(rtfFile);
            }
            catch (FreeLimitReachedException ex)
            {
                Console.WriteLine("@@@ " + ex.Message);
            }

            Console.WriteLine("@@@ Paragraphs: " + document.GetChildElements(true, ElementType.Paragraph).Count());

            foreach (Paragraph paragraph in document.GetChildElements(true, ElementType.Paragraph))
            {
                paragraphs.Add(paragraph.Content.ToString());

                //Console.WriteLine("@@@@ Paragraph: " + paragraph.Content.ToString());
            }

            return paragraphs;
        }

    }
}
