namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class InterlockDoor
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int DoorID { get; set; }

        public int OutputID { get; set; }

        public int isDisabled { get; set; }

        public int PhysicalAddress { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual InterlockSetting InterlockSetting { get; set; }
    }
}
