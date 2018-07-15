namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("LiftHLIProtocol")]
    public partial class LiftHLIProtocol
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Protocol { get; set; }

        public int DisplayMessageID { get; set; }
    }
}
