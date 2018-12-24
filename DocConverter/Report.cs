using System;
using System.Collections.Generic;
using System.Text;
using Xceed.Words.NET;
using System.Linq;

namespace DocConverter
{
    public class Report
    {
        #region Sections
        public const string Header = "EVA-PCD FUNCTIONAL PROFILE";
        public const string SingleDoseEffectHeader = "SINGLE DRUG DOSE EFFECT ANALYSIS";
        public const string MultiDoseEffectHeader = "MULTIPLE DRUG DOSE EFFECT ANALYSIS";
        public const string InterpretationHeader = "INTERPRETATION:";
        public const string ExVivoHeader = "Ex Vivo best regimen (EVBR®) would be";
        #endregion Sections

        #region Members
        private string _patient;
        private string _dx;
        private string _priorRx;
        private string _physician;
        private string _assayDate;
        private string _assayQuality;
        private string _reportDate;
        private string _specimenNumber;

        public string Patient { get => _patient; }
        public string Dx { get => _dx; }
        public string PriorRx { get => _priorRx; }
        public string Physician { get => _physician; }
        public string AssayDate { get => _assayDate; }
        public string AssayQuality { get => _assayQuality; }
        public string ReportDate { get => _reportDate; }
        public string SpecimenNumber { get => _specimenNumber; }

        public List<DrugEffect> DrugEffects { get; private set; } = new List<DrugEffect>();
        public List<MultiDrugEffect> MultiDrugEffects { get; private set; } = new List<MultiDrugEffect>();
        public IList<string> Interpretation { get; private set; } = new List<string>();
        public IList<string> Signature { get; private set; } = new List<string>();
        #endregion Members

        #region Parser
        public void Parse(IList<string> paragraphs)
        {
            var currentParagraph = paragraphs.GetEnumerator();

            VerifyHeader(currentParagraph);

            ParseHeaderInfo(currentParagraph);

            ParseSingleDoseEffect(currentParagraph);

            ParseMultiDoseEffect(currentParagraph);

            ParseExVivo(currentParagraph);

            ParseSignature(currentParagraph);
        }

        private void ParseSingleDoseEffect(IEnumerator<string> currentParagraph)
        {
            if (currentParagraph.Current.Trim().Equals(SingleDoseEffectHeader))
            {

                currentParagraph.MoveNext(); // Eat the header
                currentParagraph.MoveNext();
                var line = currentParagraph.Current.Trim();
                DrugEffect previousDrugEffect = null;

                while (!line.Equals(MultiDoseEffectHeader) && !line.Equals(InterpretationHeader))
                {
                    if (line.Contains("\t"))
                    {
                        var drugEffect = new DrugEffect(line);
                        DrugEffects.Add(drugEffect);
                        previousDrugEffect = drugEffect;
                    }
                    else
                    {
                        Console.WriteLine("*S* Before: {0}", previousDrugEffect.Drug);
                        previousDrugEffect.ExtendDrugName(line);
                        Console.WriteLine("*S*  After: {0}", previousDrugEffect.Drug);
                    }
                    currentParagraph.MoveNext();
                    line = currentParagraph.Current.Trim();
                }
            }
        }

        private void ParseMultiDoseEffect(IEnumerator<string> currentParagraph)
        {
            var line = currentParagraph.Current.Trim();

            if (line.Equals(MultiDoseEffectHeader))
            {
                DrugEffect previousDrugEffect = null;
                currentParagraph.MoveNext(); // Eat the header
                currentParagraph.MoveNext();
                line = currentParagraph.Current.Trim();

                while (!line.Equals(InterpretationHeader))
                {
                    if (line.Contains("\t"))
                    {
                        var multiDrugEffect = new MultiDrugEffect(line);
                        MultiDrugEffects.Add(multiDrugEffect);
                        previousDrugEffect = multiDrugEffect;
                    }
                    else
                    {
                        //Console.WriteLine("*M* Before: {0}", previousDrugEffect.Drug);
                        previousDrugEffect.ExtendDrugName(line);
                        //Console.WriteLine("*M*  After: {0}", previousDrugEffect.Drug);
                    }
                    currentParagraph.MoveNext();
                    line = currentParagraph.Current.Trim();
                }
            }
        }

        private string ParseExVivo(IEnumerator<string> currentParagraph)
        {
            string line;
            currentParagraph.MoveNext();
            line = currentParagraph.Current.Trim();
            while (!line.Equals(ExVivoHeader))
            {
                Interpretation.Add(line);
                currentParagraph.MoveNext();
                line = currentParagraph.Current.Trim();
            }

            return line;
        }

        private void ParseSignature(IEnumerator<string> currentParagraph)
        {
            while (currentParagraph.MoveNext())
            {
                var line = currentParagraph.Current.Trim();
                {
                    Signature.Add(line);
                }
            }
        }

        private void ParseHeaderInfo(IEnumerator<string> currentParagraph)
        {
            ParseLine(currentParagraph.Current.Trim(), out _patient, out _assayDate);

            currentParagraph.MoveNext();
            ParseLine(currentParagraph.Current.Trim(), out _dx, out _assayQuality);

            currentParagraph.MoveNext();
            ParseLine(currentParagraph.Current.Trim(), out _priorRx, out _reportDate);

            currentParagraph.MoveNext();
            ParseLine(currentParagraph.Current.Trim(), out _physician, out _specimenNumber);

            currentParagraph.MoveNext();
        }

        private void VerifyHeader(IEnumerator<string> currentParagraph)
        {
            currentParagraph.MoveNext();

            if (!currentParagraph.Current.Trim().Equals(Header))
            {
                Console.WriteLine("### !!!! Wrong Header: " + currentParagraph.Current);
            }

            currentParagraph.MoveNext();
        }

        private void ParseLine(string paragraph, out string leftValue, out string rightValue)
        {
            var items = paragraph.Split('\t');

            leftValue = items[1];
            rightValue = items[3];
        }

        public string TextReport()
        {
            var report = new StringBuilder();
            report.AppendLine(SpecLine("Patient", Patient, "Assay Date", AssayDate));
            report.AppendLine(SpecLine("Dx", Dx, "Assay Quality", AssayQuality));
            report.AppendLine(SpecLine("Prior Rx", PriorRx, "Report Date", ReportDate));
            report.AppendLine(SpecLine("Physician", Physician, "Specimen Number", SpecimenNumber));

            if (DrugEffects.Count > 0)
            {
                report.AppendLine("###" + SingleDoseEffectHeader);
                foreach (var drugEffect in DrugEffects)
                {
                    report.AppendLine(drugEffect.TextReportLine());
                }
            }

            if (MultiDrugEffects.Count > 0)
            {
                report.AppendLine("###" + MultiDoseEffectHeader);
                foreach (var drugEffect in MultiDrugEffects)
                {
                    report.AppendLine(drugEffect.TextReportLine());
                }
            }

            report.AppendLine("###" + ExVivoHeader);
            foreach (var line in Interpretation)
            {
                report.AppendLine(line);
            }

            report.AppendLine("### Signature");
            foreach (var line in Signature)
            {
                report.AppendLine(line);
            }

            return report.ToString();
        }

        private string SpecLine(string leftName, string leftValue, string rightName, string rightValue)
        {
            var outputLine = string.Format("{0}: {1}\t{2}: {3}", leftName, leftValue, rightName, rightValue);
            return outputLine;
        }
        #endregion Parse

        #region DocX Writer
        internal void DocXWrite(string docxOutputFile)
        {
            // Create a new document.
            using (var document = DocX.Create(docxOutputFile))
            {
                document.MarginLeft = 36;
                document.MarginRight = 36;


                // Add a title
                document.InsertParagraph(Header).Font("Arial").FontSize(18).Bold().UnderlineStyle(UnderlineStyle.thick).Alignment = Alignment.left;

                WriteHeaderInfo(document);

                WriteAnalysisHeader(document);

                if (DrugEffects.Count > 0)
                {
                    WriteDrugEffects(document, SingleDoseEffectHeader, DrugEffects);
                }
                if (MultiDrugEffects.Count > 0)
                {
                    WriteDrugEffects(document, MultiDoseEffectHeader, MultiDrugEffects.ToList<DrugEffect>()); //  .Cast<IDrugEffect>().ToList());
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

            foreach (var interpretationLine in Interpretation)
            {
                document.InsertParagraph(interpretationLine).Font("Arial").FontSize(10);
            }
        }

        private void WriteSignature(DocX document)
        {
            document.InsertParagraph().SpacingAfter(24);
            foreach (var signatureLine in Signature)
            {
                document.InsertParagraph(signatureLine).Font("Arial").FontSize(12);
            }
        }

        private void WriteDrugEffects(DocX document, string header, List<DrugEffect> drugEffects)
        {
            document.InsertParagraph(header).Font("Arial").FontSize(14).Bold().Alignment = Alignment.center;

            foreach (var drugEffect in drugEffects)
            {
                var paragraph = document.InsertParagraph(drugEffect.Drug)
                    .FontSize(10d)
                    .InsertTabStopPosition(Alignment.left, 225)
                    .InsertTabStopPosition(Alignment.left, 300)
                    .InsertTabStopPosition(Alignment.left, 375);
                paragraph.Append("\t" + drugEffect.Interpretation);
                paragraph.Append("\t");
                if(drugEffect is MultiDrugEffect)
                {
                    paragraph.Append((drugEffect as MultiDrugEffect).Synergy);
                }
                paragraph.Append("\tHigher").Italic();
            }
        }

        private void WriteAnalysisHeader(DocX document)
        {
            document.InsertParagraph().SpacingAfter(20);

            document.InsertParagraph("\tEx Vivo\tEx Vivo\tEx Vivo")
                .FontSize(10d).Bold()
                    .InsertTabStopPosition(Alignment.left, 225)
                    .InsertTabStopPosition(Alignment.left, 300)
                    .InsertTabStopPosition(Alignment.left, 375);
            document.InsertParagraph("\tActivity\tSynergy\tInterpretation")
                .FontSize(10d).Bold()
                    .InsertTabStopPosition(Alignment.left, 225)
                    .InsertTabStopPosition(Alignment.left, 300)
                    .InsertTabStopPosition(Alignment.left, 375);
            document.InsertParagraph("Drug\t\t\t")
                .Font("Times Roman New").FontSize(12).Bold()
                    .InsertTabStopPosition(Alignment.left, 225)
                    .InsertTabStopPosition(Alignment.left, 300)
                    .InsertTabStopPosition(Alignment.left, 375)
                .Append("Response Expectation").Font("Times Roman New");
            document.InsertParagraph("\tCompared with Database")
                .Font("Times Roman New").FontSize(10d)
                .InsertTabStopPosition(Alignment.left, 375);
            document.InsertParagraph("▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬▬")
                .FontSize(10d).Bold();
        }

        private void WriteHeaderInfo(DocX document)
        {
            document.InsertParagraph().SpacingAfter(30);

            WriteSpecLine(document, "Patient:", Patient, "Assay Date:", AssayDate);
            WriteSpecLine(document, "Dx:", Dx, "Assay Quality:", AssayQuality);
            WriteSpecLine(document, "Prior Rx:", PriorRx, "Report Date:", ReportDate);
            WriteSpecLine(document, "Physician:", Physician, "Specimen Number:", SpecimenNumber);
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
