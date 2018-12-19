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
        const string SingleDoseEffect = "SINGLE DRUG DOSE EFFECT ANALYSIS";
        const string Multi = "MULTIPLE DRUG DOSE EFFECT ANALYSIS";
        private string _patient;
        private string _dx;
        private string _priorRx;
        private string _physician;
        private string _assayDate;
        private string _assayQuality;
        private string _reportDate;
        private string _specimenNumber;

        string Patient { get => _patient; set => _patient = value; }
        string Dx { get => _dx; set => _dx = value; }
        string PriorRx { get => _priorRx; set => _priorRx = value; }
        string Physician { get => _physician; set => _physician = value; }
        string AssayDate { get => _assayDate; set => _assayDate = value; }
        string AssayQuality { get => _assayQuality; set => _assayQuality = value; }
        string ReportDate { get => _reportDate; set => _reportDate = value; }
        string SpecimenNumber { get => _specimenNumber; set => _specimenNumber = value; }

        IList<DrugEffect> DrugEffects { get; set; }
        IList<MultiDrugEffect> MultiDrugEffects { get; set; }

        string Interpretation { get; set; }

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

        }

        private void ParseLine(string paragraph, out string leftValue, out string rightValue)
        {
            var items = paragraph.Split('\t');

            leftValue = items[1];
            rightValue = items[3];
        }

        public override string ToString()
        {
            var thisString = new StringBuilder();
            thisString.AppendLine(SpecLine("Patient", Patient, "Assay Date", AssayDate));
            thisString.AppendLine(SpecLine("Dx", Dx, "Assay Quality", AssayQuality));
            thisString.AppendLine(SpecLine("Prior Rx", PriorRx, "Report Date", ReportDate));
            thisString.AppendLine(SpecLine("Physician", Physician, "Specimen Number", SpecimenNumber));

            return thisString.ToString();
        }

        private string SpecLine(string leftName, string leftValue, string rightName, string rightValue)
        {
            var outputLine = string.Format("{0}: {1}\t{2}: {3}", leftName, leftValue, rightName, rightValue);
            return outputLine;
        }
    }
}
