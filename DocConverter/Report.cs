using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocConverter
{
    public class Report
    {
        string Patient { get; set; }
        string Dx { get; set; }
        string PriorRx { get; set; }
        string Physician { get; set; }
        string AssayDate { get; set; }
        string AssayQuality { get; set; }
        string ReportDate { get; set; }
        string SpecimenNumber { get; set; }

        IList<DrugEffect> DrugEffects { get; set; }
        IList<MultiDrugEffect> MultiDrugEffects { get; set; }

        string Interpretation { get; set; }
    }
}
