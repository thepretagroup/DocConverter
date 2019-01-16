namespace DocConverter
{
    public class MultiDrugEffect : DrugEffect
    {
        public string Ratio { get; protected set; }
        public string Synergy { get; protected set; }

        public MultiDrugEffect(string drug, string ratio, string ic50, string units, string interpretation, string synergy) :
                    base(drug, ic50, units, interpretation)
        {
            Ratio = ratio;
            Synergy = synergy;
        }

        public override string TextReportLine()
        {
            return string.Join("|", Drug, Ratio, IC50, Units, Interpretation, Synergy);
        }
    }
}
