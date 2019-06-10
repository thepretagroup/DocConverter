using System.Collections.Generic;

namespace DocConverter
{
    public class MultiDrugEffect : DrugEffect
    {
        public class ExVivoSynergyFactor
        {
            public string Synergy { get; set; }
            public int Rank { get; set; }

            public ExVivoSynergyFactor(string synergy, int rank)
            {
                Synergy = synergy;
                Rank = rank;
            }
        }

        protected static readonly ExVivoSynergyFactor SynergySynergy = new ExVivoSynergyFactor("Synergy", 1);
        protected static readonly ExVivoSynergyFactor PartialSynergy = new ExVivoSynergyFactor("Partial Synergy", 2);
        protected static readonly ExVivoSynergyFactor MixedSynergy = new ExVivoSynergyFactor("Mixed Synergy", 3);
        protected static readonly ExVivoSynergyFactor NoSynergy = new ExVivoSynergyFactor("No Synergy", 4);
        protected static readonly ExVivoSynergyFactor NASynergy = new ExVivoSynergyFactor("N/A", 5);

        protected readonly Dictionary<string, ExVivoSynergyFactor> SynergyyMap = new Dictionary<string, ExVivoSynergyFactor>
        {
            { "Synergy", SynergySynergy },
            { "Partial Synergy", PartialSynergy },
            { "Mixed", MixedSynergy},
            { "Antagonism", NoSynergy },
            { "No Synergy", NoSynergy },
            { "N/A",NASynergy },
        };

        public string Ratio { get; protected set; }
        public string Synergy { get; protected set; }

        public MultiDrugEffect(string drug, string ratio, string ic50, string units, string interpretation, string synergy) :
                    base(drug, ic50, units, interpretation)
        {
            Ratio = ratio;
            Synergy = synergy;
        }

        public ExVivoSynergyFactor ExVivoSynergy
        {
            get
            {
                if (string.IsNullOrEmpty(Synergy) || !SynergyyMap.ContainsKey(Synergy))
                {
                    return NASynergy;
                }
                return SynergyyMap[Synergy];
            }
        }

        public override string TextReportLine()
        {
            return string.Join("|", Drug, Ratio, IC50, Units, Interpretation, Synergy);
        }
    }
}
