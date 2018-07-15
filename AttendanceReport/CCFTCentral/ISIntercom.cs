namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ISIntercom")]
    public partial class ISIntercom
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int IntercomSystemID { get; set; }

        [Required]
        [StringLength(64)]
        public string IntercomNumber { get; set; }

        [Required]
        [StringLength(64)]
        public string SerialNumber { get; set; }

        public int WorkstationID { get; set; }

        public int DoorID { get; set; }

        public bool AcceptCalls { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual IntercomSystem IntercomSystem { get; set; }
    }
}
