using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("CheckInInfo")]
    public class CheckInAndOutInfo
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int CheckInInfoId { get; set; }

        public string FirstName { get; set; }

        public string CNICNumber { get; set; }

        public string CardNumber { get; set; }

        public string VehicleNmuber { get; set; }

        public decimal NoOfMaleGuest { get; set; }

        public decimal NoOfFemaleGuest { get; set; }

        public decimal DurationOfStay { get; set; }

        public decimal NoOfChildren { get; set; }

        public string AreaOfVisit { get; set; }

        public string HostName { get; set; }

        public DateTime DateTimeIn { get; set; }

        public DateTime DateTimeOut { get; set; }        

        public bool CheckedIn { get; set; }

        public bool CheckInToPlant { get; set; }

        public bool CheckInToColony { get; set; }

        [ForeignKey("CardHolderInfos")]
        public int? CardHolderId { get; set; }

        public virtual CardHolderInfo CardHolderInfos { get; set; }

        [ForeignKey("Visitors")]
        public int? VisitorId { get; set; }

        public virtual VisitorCardHolder Visitors { get; set; }

        [ForeignKey("DailyCardHolders")]
        public int? DailyCardHolderId { get; set; }

        public virtual DailyCardHolder DailyCardHolders { get; set; }
    }
}
