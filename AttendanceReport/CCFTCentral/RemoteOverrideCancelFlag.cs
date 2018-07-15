namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("RemoteOverrideCancelFlag")]
    public partial class RemoteOverrideCancelFlag
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemTypeID { get; set; }

        public bool CancelFlag { get; set; }

        public virtual FTItemType FTItemType { get; set; }
    }
}
