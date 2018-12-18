namespace DocConverter
{
    public class MultiDrugEffect : DrugEffect
    {
        public string Ratio { get; set; }
        public string Synergy { get; set; }

        public MultiDrugEffect(string drugEntryLine)
        {
            var items = drugEntryLine.Split('\t');

            Drug = items[0];
            Ratio = items[1];
            IC50 = items[2];
            Units = items[3];
            Interpretation = items[4];
            Synergy = items[5];
        }
    }
}
