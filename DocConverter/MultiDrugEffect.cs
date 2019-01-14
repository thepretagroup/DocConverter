namespace DocConverter
{
    public class MultiDrugEffect : DrugEffect
    {
        public string Ratio { get; protected set; }
        public string Synergy { get; protected set; }

        public MultiDrugEffect(string[] drugEntryItems)
        {
            Drug = drugEntryItems[0].Trim();
            Ratio = drugEntryItems[1];
            IC50 = drugEntryItems[2];
            Units = drugEntryItems[3];
            Interpretation = drugEntryItems[4];
            Synergy = drugEntryItems[5];
        }

        public override string TextReportLine()
        {
            return string.Join("|", Drug, Ratio, IC50, Units, Interpretation, Synergy);
        }
    }
}
