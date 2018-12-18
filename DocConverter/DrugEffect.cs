using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocConverter
{
    public class DrugEffect
    {
        public string Drug { get; set; }
        public string IC50 { get; set; }
        public string Units { get; set; }
        public string Interpretation { get; set; }

        protected DrugEffect() {}

        public DrugEffect(string drugEntryLine)
        {
            var items = drugEntryLine.Split('\t');

            Drug = items[0];
            IC50 = items[1];
            Units = items[2];
            Interpretation = items[3];
        }
    }
}
