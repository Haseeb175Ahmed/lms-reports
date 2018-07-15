using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    [Table("CardHolder")]
    public class CardHolderInfo
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int CardHolderId { get; set; }

        public int FTItemId { get; set; }

        public string CardNumber { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string BloodGroup { get; set; }

        [ForeignKey("Cadre")]
        public int? CadreId { get; set; }

        public virtual CadreInfo Cadre { get; set; }
        
        [ForeignKey("Crew")]
        public int? CrewId { get; set; }

        public virtual CrewInfo Crew { get; set; }

        public string DateOfBirth { get; set; }

        [ForeignKey("Department")]
        public int? DepartmentId { get; set; }

        public virtual DepartmentInfo Department { get; set; }

        [ForeignKey("Designation")]
        public int? DesignationId { get; set; }

        public virtual DesignationInfo Designation { get; set; }

        [Column("ContactNumber")]
        public string EmergancyContactNumber { get; set; }

        public string CNICNumber { get; set; }

        public string PNumber { get; set; }

        [ForeignKey("Section")]
        public int? SectionId { get; set; }

        public virtual SectionInfo Section { get; set; }

        [ForeignKey("Company")]
        public int? CompanyId { get; set; }

        public virtual CompanyInfo Company { get; set; }

        public string WONumber { get; set; }

        public bool IsTemp { get; set; }

        public virtual List<BlockedPersonInfo> BlockingInfos { get; set; }

        public virtual List<CheckInAndOutInfo> CheckInInfos { get; set; }
        
        public byte[] Picture { get; set; }

        public string ConstractorInfo { get; set; }
    }
}
