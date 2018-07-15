namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("FacilityCode")]
    public partial class FacilityCode
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CodeIndex { get; set; }

        public int CodeValue { get; set; }

        public int CryptSeed { get; set; }
    }
}
