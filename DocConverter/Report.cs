using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        private IList<DrugEffect> _drugEffects = new List<DrugEffect>();
        private IList<MultiDrugEffect> _multiDrugEffects = new List<MultiDrugEffect>();
        private IList<string> _interpretation = new List<string>();

        string Patient { get => _patient; set => _patient = value; }
        string Dx { get => _dx; set => _dx = value; }
        string PriorRx { get => _priorRx; set => _priorRx = value; }
        string Physician { get => _physician; set => _physician = value; }
        string AssayDate { get => _assayDate; set => _assayDate = value; }
        string AssayQuality { get => _assayQuality; set => _assayQuality = value; }
        string ReportDate { get => _reportDate; set => _reportDate = value; }
        string SpecimenNumber { get => _specimenNumber; set => _specimenNumber = value; }

        IList<DrugEffect> DrugEffects { get => _drugEffects; set => _drugEffects = value; }
        IList<MultiDrugEffect> MultiDrugEffects { get => _multiDrugEffects; set => _multiDrugEffects = value; }

        IList<string> Interpretation { get => _interpretation; set => _interpretation = value; }

        public Report(IList<string> paragraphs)
        {
            var currentParagraph = paragraphs.GetEnumerator();
            currentParagraph.MoveNext();
            if (!currentParagraph.Current.Trim().Equals(Header))
            {
                Console.WriteLine("### Wrong Header: " + currentParagraph.Current);
            }

            currentParagraph.MoveNext();
            ParseLine(currentParagraph.Current.Trim(), out _patient, out _assayDate);

            currentParagraph.MoveNext();
            ParseLine(currentParagraph.Current.Trim(), out _dx, out _assayQuality);

            currentParagraph.MoveNext();
            ParseLine(currentParagraph.Current.Trim(), out _priorRx, out _reportDate);

            currentParagraph.MoveNext();
            ParseLine(currentParagraph.Current.Trim(), out _physician, out _specimenNumber);

            currentParagraph.MoveNext();
            if (!currentParagraph.Current.Trim().Equals(SingleDoseEffectHeader))
            {
                Console.WriteLine("### Wrong SingleDoseEffect: " + currentParagraph.Current);
            }

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
                    previousDrugEffect.Drug += " " + line;
                    Console.WriteLine("*S*  After: {0}", previousDrugEffect.Drug);
                }
                currentParagraph.MoveNext();
                line = currentParagraph.Current.Trim();
            }

            if (line.Equals(MultiDoseEffectHeader))
            {
                previousDrugEffect = null;
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
                        Console.WriteLine("*M* Before: {0}", previousDrugEffect.Drug);
                        previousDrugEffect.Drug += " " + line;
                        Console.WriteLine("*M*  After: {0}", previousDrugEffect.Drug);
                    }
                    currentParagraph.MoveNext();
                    line = currentParagraph.Current.Trim();
                }
            }

            currentParagraph.MoveNext();
            line = currentParagraph.Current.Trim();
            while (!line.Equals(ExVivoHeader))
            {
                Interpretation.Add(line);
                currentParagraph.MoveNext();
                line = currentParagraph.Current.Trim();
            }
        }

        private void ParseLine(string paragraph, out string leftValue, out string rightValue)
        {
            var items = paragraph.Split('\t');

            leftValue = items[1];
            rightValue = items[3];
        }

        public override string ToString()
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
                    report.AppendLine(drugEffect.ToString());
                }
            }


            if (MultiDrugEffects.Count > 0)
            {
                report.AppendLine("###" + MultiDoseEffectHeader);
                foreach (var drugEffect in MultiDrugEffects)
                {
                    report.AppendLine(drugEffect.ToString());
                }
            }

            report.AppendLine("###" + ExVivoHeader);
            foreach (var line in Interpretation)
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
