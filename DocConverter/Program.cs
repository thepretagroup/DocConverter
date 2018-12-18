using System;
using System.Linq;
using System.Text;
using GemBox.Document;
using GemBox.Document.Tables;
using System.Text.RegularExpressions;

namespace DocConverter
{
    class Sample
    {
        [STAThread]
        static void Main(string[] args)
        {
            // If using Professional version, put your serial key below.
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            // Continue to use the component in a Trial mode when free limit is reached.
            ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;

            RTFtoTxt();
            //CreateTableDoc();
        }

        static void RTFtoTxt()
        {
            DocumentModel document = null;

            try
            {
                //document = DocumentModel.Load(@"C:\Users\dhfra\Documents\GitHub\DocConverter\TestDocShort.rtf");
                document = DocumentModel.Load(@"C:\Users\dhfra\Documents\GitHub\DocConverter\input file.rtf");
            }
            catch (FreeLimitReachedException ex)
            {
                Console.WriteLine("@@@ " + ex.Message);
            }

            StringBuilder sb = new StringBuilder();

            Console.WriteLine("@@@ Paragraphs: " + document.GetChildElements(true, ElementType.Paragraph).Count());
            var report = new Report();

            foreach (Paragraph paragraph in document.GetChildElements(true, ElementType.Paragraph))
            {
                Console.WriteLine(" Paragraph: " + paragraph.Content.ToString());
                //foreach (Run run in paragraph.GetChildElements(true, ElementType.Run))
                //{
                //    bool isBold = run.CharacterFormat.Bold;
                //    string text = run.Text;

                //    sb.AppendFormat("{0}{1}{2}", isBold ? "<b>" : "", text, isBold ? "</b>" : "");
                //}
                //sb.AppendLine();
            }

            document.Save(@"C:\Users\dhfra\Documents\GitHub\DocConverter\myOutput.txt");
            //document.Save(@"C:\Users\dhfra\Documents\GitHub\DocConverter\TestDocShort_OUT.txt");

            //Console.WriteLine(sb.ToString());
        }

        static void CreateTableDoc()
        {
            DocumentModel document = new DocumentModel();

            int tableRowCount = 10;
            int tableColumnCount = 5;

            Table table = new Table(document);
            table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);

            for (int i = 0; i < tableRowCount; i++)
            {
                TableRow row = new TableRow(document);
                table.Rows.Add(row);

                for (int j = 0; j < tableColumnCount; j++)
                {
                    Paragraph para = new Paragraph(document, string.Format("Cell {0}-{1}", i + 1, j + 1));

                    row.Cells.Add(new TableCell(document, para));
                }
            }

            document.Sections.Add(new Section(document, table));

            document.Save(@"C:\Users\dhfra\Documents\GitHub\DocConverter\Simple Table.docx");
        }
    }
}

