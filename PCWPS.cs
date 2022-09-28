using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;


namespace TrainingApp.Models
{
    public class PCWPS
    {
        [Key]
        public int Pcwbs_Id { get; set; }
        [Required(ErrorMessage = "Required")]
        public int Job_Code_Id { get; set; }
        [Required(ErrorMessage = "Required")]
        public string Prime_Code { get; set; }
        [Required(ErrorMessage = "Required")]
        public string Pcwbs { get; set; }
        public int Parent_Pcwbs_Id { get; set; }
        public string Pcwbs_Descr { get; set; }
        public Nullable<int> Pcwbs_Level_No { get; set; }
        public Nullable<int> Priority_Num { get; set; }
        public string Plan_Shop_Subcon { get; set; }
        public string Plan_Field_Subcon { get; set; }
        public bool Active_Flg { get; set; }
        public string Remarks { get; set; }
        public System.DateTime Modified_On { get; set; }
        public System.DateTime Created_On { get; set; }
    }
}
