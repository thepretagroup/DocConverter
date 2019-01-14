namespace DocConverter
{
    public class DrugEffect
    {
        public string Drug { get; protected set; }
        public string IC50 { get; protected set; }
        public string Units { get; protected set; }
        public string Interpretation { get; protected set; }

        protected DrugEffect() {}

        public DrugEffect(string[] drugEntryItems)
        {
            //var drugEntryItems = drugEntryLine.Split('\t');

            Drug = drugEntryItems[0].Trim();
            IC50 = drugEntryItems[1];
            Units = drugEntryItems[2];
            Interpretation = drugEntryItems[3];
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
