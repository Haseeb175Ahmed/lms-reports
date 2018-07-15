using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("DailyCardHolder")]
    public class DailyCardHolder
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int DailyCardHolderId { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string CompanyName { get; set; }

        public string Cadre { get; set; }

        public string Department { get; set; }

        public string Designation { get; set; }

        public string Section { get; set; }

        [Column("ContactNumber")]
        public string EmergancyContactNumber { get; set; }

        public string CNICNumber { get; set; }

        public string WONumber { get; set; }

        public virtual List<BlockedPersonInfo> BlockingInfos { get; set; }

        public virtual List<CheckInAndOutInfo> CheckInInfos { get; set; }

        public byte[] Picture { get; set; }

        public string ClubType { get; set; }

        public string ConstractorInfo { get; set; }

        public string AreaOfWork { get; set; }
    }
}
