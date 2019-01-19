using System.Collections.Generic;

namespace DocConverter
{
    public class DrugEffect
    {
        public class ExVivoFactor
        {
            public string Activity { get; set; }
            public string Interpretation { get; set; }
            public int Rank { get; set; }

            public ExVivoFactor(string activity, string interpretation, int rank)
            {
                Activity = activity;
                Interpretation = interpretation;
                Rank = rank;
            }
        }

        protected static readonly ExVivoFactor LowerExVivo = new ExVivoFactor("Resistant", "Lower", 3);
        protected static readonly ExVivoFactor HigherExVivo = new ExVivoFactor("Sensitive", "Higher", 1);
        protected static readonly ExVivoFactor IntermediateExVivo = new ExVivoFactor("Intermediate", "Average", 2);
        protected static readonly ExVivoFactor UnknownExVivo = new ExVivoFactor("?", "?", 4);

        protected readonly Dictionary<string, ExVivoFactor> ActivityMap = new Dictionary<string, ExVivoFactor>
        {
            { "Resistant", LowerExVivo},
            { "Inactive",LowerExVivo },
            { "Sensitive", HigherExVivo },
            { "Intermediate", IntermediateExVivo },
            { "Moderately Active", IntermediateExVivo },
            { "Active", HigherExVivo },
        };

        public string Drug { get; protected set; }
        public string IC50 { get; protected set; }
        public string Units { get; protected set; }
        public string Interpretation { get; internal set; }
        public ExVivoFactor ExVivo
        {
            get
            {
                if (string.IsNullOrEmpty(Interpretation) || !ActivityMap.ContainsKey(Interpretation))
                {
                    return UnknownExVivo;
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
