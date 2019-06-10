using DocConverter.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocConverter
{
    public class Report
    {
        #region Members
        private string _patient;
        private string _dx;
        private string _priorRx;
        private List<string> _physician = new List<string>();
        private string _assayDate;
        private string _assayQuality;
        private string _reportDate;
        private string _specimenNumber;

        public string Patient { get => _patient; }
        public string Dx { get => _dx; }
        public string PriorRx { get => _priorRx; }
        public string[] Physician { get => _physician.ToArray(); }
        public string AssayDate { get => _assayDate; }
        public string AssayQuality { get => _assayQuality; }
        public string ReportDate { get => _reportDate; }
        public string SpecimenNumber { get => _specimenNumber; }

        public List<DrugEffect> DrugEffects { get; private set; } = new List<DrugEffect>();
        public List<MultiDrugEffect> MultiDrugEffects { get; private set; } = new List<MultiDrugEffect>();
        public IList<string> Signature { get; private set; } = new List<string>();
        #endregion Members

        #region Parser
        public void Parse(IList<string> paragraphs)
        {
            var currentParagraph = 0;
            try
            {
                VerifyHeader(paragraphs, ref currentParagraph);

                ParseHeaderInfo(paragraphs, ref currentParagraph);

                ParseSingleDoseEffect(paragraphs, ref currentParagraph);

                ParseMultiDoseEffect(paragraphs, ref currentParagraph);

                ParseExVivo(paragraphs, ref currentParagraph);

                ParseSignature(paragraphs, ref currentParagraph);

                if (currentParagraph < paragraphs.Count)
                {
                    SearchForMoreDoctors(paragraphs, ref currentParagraph);
                }
            }
            catch (Exception ex)
            {
                throw new ParseException("Input file parsing error in paragraph #" + currentParagraph, ex);
            }
        }

        private void SearchForMoreDoctors(IList<string> paragraphs, ref int currentParagraph)
        {
            for (currentParagraph++; currentParagraph < paragraphs.Count; currentParagraph++)
            {
                var line = paragraphs[currentParagraph].Trim();

                var items = line.Split('\t');
                if (items.Length == 4 && items[0].Equals("Physician:"))
                {
                    _physician.Add(items[1]);
                }
            }
        }

        private void ParseSingleDoseEffect(IList<string> paragraphs, ref int currentParagraph)
        {
            if (paragraphs[currentParagraph].Trim().Equals(Resources.SingleDoseEffectHeader))
            {

                currentParagraph++; // Eat the header
                var line = paragraphs[++currentParagraph].Trim();
                DrugEffect previousDrugEffect = null;

                while (!line.Equals(Resources.MultiDoseEffectHeader) && !line.Equals(Resources.InterpretationHeader))
                {
                    if (line.Contains("\t"))
                    {
                        var drugEntryItems = line.Split('\t');
                        if (drugEntryItems.Count() == 4)
                        {
                            var drugEffect = new DrugEffect(drugEntryItems[0], drugEntryItems[1], drugEntryItems[2], drugEntryItems[3]);
                            DrugEffects.Add(drugEffect);
                            previousDrugEffect = drugEffect;
                        }
                        else
                        {
                            previousDrugEffect = null;
                        }
                    }
                    else
                    {
                        previousDrugEffect.ExtendDrugName(line);
                    }
                    line = paragraphs[++currentParagraph].Trim();
                }
            }
        }

        private void ParseMultiDoseEffect(IList<string> paragraphs, ref int currentParagraph)
        {
            var line = paragraphs[currentParagraph].Trim();

            if (line.Equals(Resources.MultiDoseEffectHeader))
            {
                DrugEffect previousDrugEffect = null;
                currentParagraph++; // Eat the header
                line = paragraphs[++currentParagraph].Trim();

                while (!line.Equals(Resources.InterpretationHeader))
                {
                    if (line.Contains("\t"))
                    {
                        var drugEntryItems = line.Split('\t').ToList();
                        if (drugEntryItems.Count() == 5)  // Ratio is missing
                        {
                            drugEntryItems.Insert(1, string.Empty);
                        }

                        if (drugEntryItems.Count() == 6 && !string.IsNullOrEmpty(drugEntryItems[4]))
                        {
                            var multiDrugEffect = new MultiDrugEffect(
                                drugEntryItems[0], drugEntryItems[1], drugEntryItems[2], drugEntryItems[3], drugEntryItems[4], drugEntryItems[5]);

                            MultiDrugEffects.Add(multiDrugEffect);
                            previousDrugEffect = multiDrugEffect;
                        }
                        else
                        {
                            previousDrugEffect = null;
                        }
                    }
                    else
                    {
                        previousDrugEffect.ExtendDrugName(line);
                    }
                    line = paragraphs[++currentParagraph].Trim();
                }
            }
        }

        private string ParseExVivo(IList<string> paragraphs, ref int currentParagraph)
        {
            string line;
            line = paragraphs[++currentParagraph].Trim();
            while (!line.Contains(Resources.ExVivoHeader))
            {
                line = paragraphs[++currentParagraph].Trim();
            }

            return line;
        }

        private void ParseSignature(IList<string> paragraphs, ref int currentParagraph)
        {
            for (currentParagraph++; currentParagraph < paragraphs.Count; currentParagraph++)
            {
                var line = paragraphs[currentParagraph].Trim();
                if (line.Equals(Resources.ReportHeader))
                {
                    Console.WriteLine("New page found");
                    break;
                }
                Signature.Add(line);
            }
        }

        private void ParseHeaderInfo(IList<string> paragraphs, ref int currentParagraph)
        {
            ParseLine(paragraphs[currentParagraph].Trim(), out _patient, out _assayDate);

            ParseLine(paragraphs[++currentParagraph].Trim(), out _dx, out _assayQuality);

            ParseLine(paragraphs[++currentParagraph].Trim(), out _priorRx, out _reportDate);

            ParseLine(paragraphs[++currentParagraph].Trim(), out string physician, out _specimenNumber);
            _physician.Add(physician);

            currentParagraph++;
        }

        private void VerifyHeader(IList<string> paragraphs, ref int currentParagraph)
        {
            if (!paragraphs[currentParagraph].Trim().Equals(Resources.ReportHeader))
            {
                Console.WriteLine("### !!!! Wrong Header: " + paragraphs[currentParagraph]);
            }

            currentParagraph++;
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
            report.AppendLine(SpecLine("Physician", Physician[0], "Specimen Number", SpecimenNumber));

            if (DrugEffects.Count > 0)
            {
                report.AppendLine("###" + Resources.SingleDoseEffectHeader);
                foreach (var drugEffect in DrugEffects)
                {
                    report.AppendLine(drugEffect.TextReportLine());
                }
            }

            if (MultiDrugEffects.Count > 0)
            {
                report.AppendLine("###" + Resources.MultiDoseEffectHeader);
                foreach (var drugEffect in MultiDrugEffects)
                {
                    report.AppendLine(drugEffect.TextReportLine());
                }
            }

            report.AppendLine("###" + Resources.ExVivoHeader);
            report.AppendLine("###" + Resources.DataAnalysis);
            report.AppendLine("###" + Resources.DataAnalysisNote);

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
    }
}
