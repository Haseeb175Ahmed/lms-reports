using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("VisitingLocations")]
    public class VisitingLocations
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int VisitingLocationsId { get; set; }

        public string Location { get; set; }

        public bool IsOnPlant { get; set; }
    }

   
}
