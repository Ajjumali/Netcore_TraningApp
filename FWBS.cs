using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace TrainingApp.Models
{
    public class FWBS
    {
        [Key]
        public int Fwbs_Id { get; set; }
        public string ProjectName { get; set; }
        public string Prime_Code { get; set; }
        public string Fwbs { get; set; }
        public int Parent_Fwbs_Id { get; set; }
        public string Fwbs_Descr { get; set; }
        public Nullable<int> Fwbs_Level_No { get; set; }
        public string BQ_Unit { get; set; }
    }
}
