using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace TrainingApp.Models
{
    public class MileStone
    {
        [Key]
        public int MileStone_Id { get; set; }
        public string ProjetName { get; set; }
        public string MileStone_Code { get; set; }
        public string MileStone_Descr { get; set; }
        
        public int FKFwbs { get; set; }
        public int FKPcwbs { get; set; }
    }
}
