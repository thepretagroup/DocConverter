using System;
using System.Collections.Generic;
using System.Text;

namespace DocConverter
{
    public class Report
    {
        const string Header = "EVA-PCD FUNCTIONAL PROFILE";
        const string SingleDoseEffectHeader = "SINGLE DRUG DOSE EFFECT ANALYSIS";
        const string MultiDoseEffectHeader = "MULTIPLE DRUG DOSE EFFECT ANALYSIS";
        const string InterpretationHeader = "INTERPRETATION:";
        const string ExVivoHeader = "Ex Vivo best regimen (EVBR®) would be";

        private string _patient;
        private string _dx;
        private string _priorRx;
        private string _physician;
        private string _assayDate;
        private string _assayQuality;
        private string _reportDate;
        private string _specimenNumber;

        public string Patient { get => _patient;}
        public string Dx { get => _dx;}
        public string PriorRx { get => _priorRx;}
        public string Physician { get => _physician;}
        public string AssayDate { get => _assayDate;}
        public string AssayQuality { get => _assayQuality;}
        public string ReportDate { get => _reportDate;}
        public string SpecimenNumber { get => _specimenNumber;}

        public IList<DrugEffect> DrugEffects { get; private set; } = new List<DrugEffect>();
        public IList<MultiDrugEffect> MultiDrugEffects { get; private set; } = new List<MultiDrugEffect>();
        public IList<string> Interpretation { get; private set; } = new List<string>();
        public IList<string> Signature { get; private set; } = new List<string>();

        public Report(IList<string> paragraphs)
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
    }
}
