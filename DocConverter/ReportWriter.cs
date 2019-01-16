using System;
using System.Collections.Generic;
using System.Text;
using Xceed.Words.NET;
using System.Linq;

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

                WriteAnalysisHeader(document);

                if (Report.DrugEffects.Count > 0)
                {
                    WriteDrugEffects(document, Report.SingleDoseEffectHeader, Report.DrugEffects);
                }
                if (Report.MultiDrugEffects.Count > 0)
                {
                    WriteDrugEffects(document, Report.MultiDoseEffectHeader, Report.MultiDrugEffects.ToList<DrugEffect>()); //  .Cast<IDrugEffect>().ToList());
                }

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

        private void WriteDrugEffects(DocX document, string header, List<DrugEffect> drugEffects)
        {
            document.InsertParagraph(header).Font("Arial").FontSize(14).Bold().Alignment = Alignment.center;

            foreach (var drugEffect in drugEffects.OrderBy(de => de.Drug))
            {
                var paragraph = document.InsertParagraph(drugEffect.Drug)
                    .FontSize(10d)
                    .InsertTabStopPosition(Alignment.left, 225)
                    .InsertTabStopPosition(Alignment.left, 315)
                    .InsertTabStopPosition(Alignment.left, 390);
                paragraph.Append("\t" + drugEffect.Interpretation);
                paragraph.Append("\t");
                if (drugEffect is MultiDrugEffect)
                {
                    paragraph.Append((drugEffect as MultiDrugEffect).Synergy);
                }
                paragraph.Append("\t" + drugEffect.ExVivoInterpretation);
            }
        }

        private void WriteAnalysisHeader(DocX document)
        {
            document.InsertParagraph().SpacingAfter(20);

            document.InsertParagraph("\tEx Vivo\tEx Vivo\tEx Vivo")
                .FontSize(10d).Bold()
                    .InsertTabStopPosition(Alignment.left, 225)
                    .InsertTabStopPosition(Alignment.left, 315)
                    .InsertTabStopPosition(Alignment.left, 390);
            document.InsertParagraph("\tActivity\tSynergy\tInterpretation")
                .FontSize(10d).Bold()
                    .InsertTabStopPosition(Alignment.left, 225)
                    .InsertTabStopPosition(Alignment.left, 315)
                    .InsertTabStopPosition(Alignment.left, 390);
            document.InsertParagraph("Drug\t\t\t")
                .Font("Times Roman New").FontSize(12).Bold()
                    .InsertTabStopPosition(Alignment.left, 225)
                    .InsertTabStopPosition(Alignment.left, 315)
                    .InsertTabStopPosition(Alignment.left, 390)
                .Append("Response Expectation").Font("Times Roman New");
            document.InsertParagraph("\tCompared with Database")
                .Font("Times Roman New").FontSize(10d)
                .InsertTabStopPosition(Alignment.left, 390);
            document.InsertParagraph("▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬")
                .FontSize(10d).Bold();
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