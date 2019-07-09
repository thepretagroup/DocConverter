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

        protected static readonly ExVivoFactor HigherExVivo = new ExVivoFactor("Sensitive", "Higher", 1);
        protected static readonly ExVivoFactor IntermediateExVivo = new ExVivoFactor("Intermediate", "Average", 2);
        protected static readonly ExVivoFactor ResistantExVivo = new ExVivoFactor("Resistant", "Lower", 3);
        protected static readonly ExVivoFactor ActiveExVivo = new ExVivoFactor("Active", "Higher", 4);
        protected static readonly ExVivoFactor ModeratelyActiveExVivo = new ExVivoFactor("Moderately Active", "Average", 5);
        protected static readonly ExVivoFactor LowerExVivo = new ExVivoFactor("Inactive", "Lower", 6);
        protected static readonly ExVivoFactor UnknownExVivo = new ExVivoFactor("?", "?", 999);

        protected readonly Dictionary<string, ExVivoFactor> ActivityMap = new Dictionary<string, ExVivoFactor>
        {
            { "Sensitive", HigherExVivo },
            { "Intermediate", IntermediateExVivo },
            { "Resistant", ResistantExVivo},
            { "Active", ActiveExVivo },
            { "Moderately Active", ModeratelyActiveExVivo },
            { "Inactive",LowerExVivo },
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
                    System.Console.WriteLine("UnknownExVivo for '{0}', interpretation = '{1}'", Drug, Interpretation);
                    return UnknownExVivo;
                }
                return ActivityMap[Interpretation];
            }
        }

        protected DrugEffect() { }

        public DrugEffect(string drug, string ic50, string units, string interpretation)
        {
            Drug = drug.Trim();
            IC50 = ic50;
            Units = units;
            // Fix possible truncated Moderately Active
            if (interpretation != null && interpretation.StartsWith("Moderately"))
            {
                interpretation = ModeratelyActiveExVivo.Activity;
            }
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
