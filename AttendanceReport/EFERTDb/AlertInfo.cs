using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("AlertInfo")]
    public class AlertInfo
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        public string CNICNumber { get; set; }

        public bool DisableAlert { get; set; }

        public DateTime DisableAlertDate { get; set; }

        public DateTime EnableAlertDate { get; set; }
    }
}
