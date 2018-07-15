using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("BlockedPersonInfo")]
    public class BlockedPersonInfo
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int BlockedPersonInfoId { get; set; }

        public string CNICNumber { get; set; }

        public string BlockedBy { get; set; }

        public string BlockedReason { get; set; }

        public string UnBlockedBy { get; set; }

        public string UnBlockedReason { get; set; }

        public bool Blocked { get; set; }

        public DateTime BlockedTime { get; set; }

        public DateTime UnBlockTime { get; set; }

        public bool BlockedInPlant { get; set; }

        public bool BlockedInColony { get; set; }

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
