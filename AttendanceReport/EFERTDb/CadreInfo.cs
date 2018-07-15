using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("Cadre")]
    public class CadreInfo
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int CadreId { get; set; }

        public string CadreName { get; set; }

        public virtual List<CardHolderInfo> CardHolders { get; set; }
    }
}
