using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("Visitor")]
    public class VisitorCardHolder
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int VisitorId { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string Address { get; set; }

        public string Gender { get; set; }

        public string PostCode { get; set; }

        public string City { get; set; }
        
        public string CNICNumber { get; set; }

        public string State { get; set; }

        public string CompanyName { get; set; }

        public string ContactNo { get; set; }

        [Column("ContantPerson")]
        public string EmergencyContantPerson { get; set; }

        [Column("EmergencytNumber")]
        public string EmergencyContantPersonNumber { get; set; }

        public string VisitorType { get; set; }

        public virtual List<BlockedPersonInfo> BlockingInfos { get; set; }

        public virtual List<CheckInAndOutInfo> CheckInInfos { get; set; }

        public bool IsOnPlant { get; set; }
                
        public byte[] Picture { get; set; }
    
        public string SchoolName { get; set; }

        public string VisitorInfo { get; set; }
    }
}
