using System;
using System.Collections.Generic;
using System.Linq;
using Xceed.Words.NET;

namespace DocConverter
{
    public class ReportWriter
    {
        private Report Report { get; set; }

        public ReportWriter(Report report)
        {
            Report = report;
        }

        #region DocX Writer
        public void CreateDocX(string docxOutputFile)
        {
            // Create a new document.
            using (var document = DocX.Create(docxOutputFile))
            {
                document.MarginLeft = 36;
                document.MarginRight = 36;

                // Add a title
                document.InsertParagraph(Report.Header).Font("Arial").FontSize(18).Bold().UnderlineStyle(UnderlineStyle.thick).Alignment = Alignment.left;

                WriteHeaderInfo(document);
                document.InsertParagraph().SpacingAfter(20);

                var table = CreateAnalysisTable(document);

                if (Report.DrugEffects.Count > 0)
                {
                    WriteDrugEffects(table, Report.SingleDoseEffectHeader, Report.DrugEffects);
                }
                if (Report.MultiDrugEffects.Count > 0)
                {
                    WriteDrugEffects(table, Report.MultiDoseEffectHeader, Report.MultiDrugEffects.ToList<DrugEffect>()); //  .Cast<IDrugEffect>().ToList());
                }

                document.InsertTable(table);

                WriteInterpretation(document);

                WriteExVivoBest(document);

                WriteSignature(document);

                // Save this document to disk.
                document.Save();
            }
        }

        private void WriteExVivoBest(DocX document)
        {
            var exVivoMsg = "Ex Vivo best regimen (EVBR®) would be Irinotecan +/ -Cisplatin.";

            document.InsertParagraph().SpacingAfter(24);

            document.InsertParagraph(exVivoMsg).Font("Arial").FontSize(12).Bold().Alignment = Alignment.left;
        }

        private void WriteInterpretation(DocX document)
        {
            document.InsertParagraph().SpacingAfter(24);

            document.InsertParagraph("DATA ANALYSIS:").Font("Arial").FontSize(14).Bold().Alignment = Alignment.left;

            foreach (var interpretationLine in Report.Interpretation)
            {
                document.InsertParagraph(interpretationLine).Font("Arial").FontSize(10);
            }
        }

        private void WriteSignature(DocX document)
        {
            document.InsertParagraph().SpacingAfter(24);
            foreach (var signatureLine in Report.Signature)
            {
                document.InsertParagraph(signatureLine).Font("Arial").FontSize(12);
            }
        }

        private void WriteDrugEffects(Table table, string header, List<DrugEffect> drugEffects)
        {
            var row = table.InsertRow();
            row.MergeCells(0, 3);
            row.Cells[0].Paragraphs[0].Append(header).Font("Arial").FontSize(14).Bold().Alignment = Alignment.center;

            foreach (var drugEffect in drugEffects.OrderBy(de => de.Drug))
            {
                row = table.InsertRow();
                row.Cells[0].Paragraphs[0].Append(drugEffect.Drug).FontSize(10d);
                row.Cells[1].Paragraphs[0].Append(drugEffect.Interpretation).FontSize(10d);
                if (drugEffect is MultiDrugEffect)
                {
                    row.Cells[2].Paragraphs[0].Append((drugEffect as MultiDrugEffect).Synergy).FontSize(10d);
                }
                row.Cells[3].Paragraphs[0].Append(drugEffect.ExVivoInterpretation).FontSize(10d);
            }
        }

        private Table CreateAnalysisTable(DocX document) {
            var table = document.AddTable(4, 4);

            table.Design = TableDesign.None;
            table.AutoFit = AutoFit.ColumnWidth;
            table.SetWidths(new[] { 225f, 90f, 75f, 120f });

            table.Rows[0].Cells[1].Paragraphs[0].Append("Ex Vivo").FontSize(10d).Bold();
            table.Rows[0].Cells[2].Paragraphs[0].Append("Ex Vivo").FontSize(10d).Bold();
            table.Rows[0].Cells[3].Paragraphs[0].Append("Ex Vivo").FontSize(10d).Bold();
            table.Rows[1].Cells[1].Paragraphs[0].Append("Activity").FontSize(10d).Bold();
            table.Rows[1].Cells[2].Paragraphs[0].Append("Synergy").FontSize(10d).Bold();
            table.Rows[1].Cells[3].Paragraphs[0].Append("Interpretation").FontSize(10d).Bold();
            table.Rows[2].Cells[0].Paragraphs[0].Append("Drug").Font("Times Roman New").FontSize(12).Bold();
            table.Rows[2].Cells[3].Paragraphs[0].Append("Response Expectation").Font("Times Roman New").FontSize(10d);
            table.Rows[3].Cells[3].Paragraphs[0].Append("Compared with Database").Font("Times Roman New").FontSize(10d);

            var border = new Border(BorderStyle.Tcbs_thick, BorderSize.six, 0, System.Drawing.Color.Black);
            foreach (var cell in table.Rows[3].Cells)
            {
                cell.SetBorder(TableCellBorderType.Bottom, border);
            }

            return table;
        }

        private void WriteHeaderInfo(DocX document)
        {
            document.InsertParagraph().SpacingAfter(30);

            WriteSpecLine(document, "Patient:", Report.Patient, "Assay Date:", Report.AssayDate);
            WriteSpecLine(document, "Dx:", Report.Dx, "Assay Quality:", "" /* Report.AssayQuality */);
            WriteSpecLine(document, "Prior Rx:", Report.PriorRx, "Report Date:", Report.ReportDate);
            WriteSpecLine(document, "Physician:", Report.Physician, "Specimen Number:", Report.SpecimenNumber);
        }

        private void WriteSpecLine(DocX document, string leftName, string leftValue, string rightName, string rightValue)
        {
            var outputLine = string.Format("{0}\t{1}\t{2}\t{3}", leftName, leftValue, rightName, rightValue);
            Console.WriteLine("@@@@@@" + outputLine);

            var paragraph = document.InsertParagraph(outputLine)
                .FontSize(10d)
                .Bold()
                .InsertTabStopPosition(Alignment.left, 100)
                .InsertTabStopPosition(Alignment.left, 300)
                .InsertTabStopPosition(Alignment.left, 400);
        }
        #endregion
    }
}