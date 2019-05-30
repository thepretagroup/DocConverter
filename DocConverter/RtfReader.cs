using System.Collections.Generic;
using System.IO;

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
            Stream fileStream = null;
            try
            {
                fileStream = new FileStream(rtfFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using (var textReader = new StreamReader(fileStream))
                {
                    fileStream = null; // Avoid CA2202 warning - https://docs.microsoft.com/en-us/visualstudio/code-quality/ca2202-do-not-dispose-objects-multiple-times?view=vs-2017
                    // Based on example in https://social.msdn.microsoft.com/Forums/vstudio/en-US/6e56af9b-d7d3-49f3-9ec4-80edde3fe54b/reading-modifying-rtf-files?forum=csharpgeneral
                    System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();
                    var rtfText = textReader.ReadToEnd();

                    rtBox.Rtf = rtfText;
                    string plainText = rtBox.Text;

                    var paragraphs = plainText.Split('\n');
                    return paragraphs;
                }
            }
            finally
            {
                if (fileStream != null)
                    fileStream.Dispose();
            }
        }
    }
}
