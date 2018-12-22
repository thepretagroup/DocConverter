namespace DocConverter
{
    public class MultiDrugEffect : DrugEffect
    {
        public string Ratio { get; protected set; }
        public string Synergy { get; protected set; }

        public MultiDrugEffect(string drugEntryLine)
        {
            var items = drugEntryLine.Split('\t');

            Drug = items[0].Trim();
            Ratio = items[1];
            IC50 = items[2];
            Units = items[3];
            Interpretation = items[4];
            Synergy = items[5];
        }

        public override string TextReportLine()
        {
            return string.Join("|", Drug, Ratio, IC50, Units, Interpretation, Synergy);
        }
    }
}
