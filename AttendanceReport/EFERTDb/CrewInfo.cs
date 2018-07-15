using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("Crew")]
    public class CrewInfo
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int CrewId { get; set; }

        public string CrewName { get; set; }

        public virtual List<CardHolderInfo> CardHolders { get; set; }
    }
}
