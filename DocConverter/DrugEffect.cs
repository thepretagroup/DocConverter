using System.Collections.Generic;

namespace DocConverter
{
    public class DrugEffect
    {
        private readonly Dictionary<string, string> ActivityMap = new Dictionary<string, string>
        {
            { "Resistant","Lower" },
            { "Inactive","Lower" },
            { "Sensitive", "Higher" },
            { "Intermediate", "Average" },
            { "Moderately Active", "Average" },
            { "Active", "** Mapping ???" },  // TODO:  What is this mapped to?
        };

        public string Drug { get; protected set; }
        public string IC50 { get; protected set; }
        public string Units { get; protected set; }
        public string Interpretation { get; internal set; }
        public string ExVivoInterpretation
        {
            get
            {
                if (string.IsNullOrEmpty(Interpretation) || !ActivityMap.ContainsKey(Interpretation))
                {
                    return "?";   // TODO:  What is the correct value? should this be String.Empty ?
                }

                return ActivityMap[Interpretation];
            }
        }

        protected DrugEffect() { }

        public DrugEffect(string drug, string ic50, string units, string interpretation)
        {
            //var drugEntryItems = drugEntryLine.Split('\t');

            Drug = drug.Trim();
            IC50 = ic50;
            Units = units;
            Interpretation = interpretation;
        }

        public void ExtendDrugName(string line)
        {
            Drug += " " + line;
        }

        public virtual string TextReportLine()
        {
            return string.Join("|", Drug, IC50, Units, Interpretation);
        }
    }
}
