namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("AperioDoor")]
    public partial class AperioDoor
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? EntryZoneID { get; set; }

        public int? CardTypeID { get; set; }

        public int CardSwipeWait { get; set; }

        public int DOTLWarnTime { get; set; }

        public int DOTLAlarmTime { get; set; }

        public int DoubleCardOverrideMode { get; set; }

        public virtual AccessZone AccessZone { get; set; }

        public virtual CardType CardType { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }
    }
}
