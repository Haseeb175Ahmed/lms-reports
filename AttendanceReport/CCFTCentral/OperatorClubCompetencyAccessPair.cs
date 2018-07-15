namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("OperatorClubCompetencyAccessPair")]
    public partial class OperatorClubCompetencyAccessPair
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int OperatorClubID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CompetencyID { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AccessLevel { get; set; }

        [Key]
        [Column(Order = 3)]
        public Guid GlobalID { get; set; }

        public virtual Competency Competency { get; set; }

        public virtual OperatorClub OperatorClub { get; set; }
    }
}
