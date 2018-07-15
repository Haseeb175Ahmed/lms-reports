namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("OperatorClubWorkstationPair")]
    public partial class OperatorClubWorkstationPair
    {
        public int OperatorClubID { get; set; }

        public int WorkstationID { get; set; }

        [Key]
        public Guid GlobalID { get; set; }

        public virtual FTItem FTItem { get; set; }

        public virtual FTItem FTItem1 { get; set; }
    }
}
