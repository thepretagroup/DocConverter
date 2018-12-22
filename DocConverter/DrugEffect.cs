namespace DocConverter
{
    public class DrugEffect
    {
        public string Drug { get; protected set; }
        public string IC50 { get; protected set; }
        public string Units { get; protected set; }
        public string Interpretation { get; protected set; }

        protected DrugEffect() {}

        public DrugEffect(string drugEntryLine)
        {
            var items = drugEntryLine.Split('\t');

            Drug = items[0].Trim();
            IC50 = items[1];
            Units = items[2];
            Interpretation = items[3];
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
