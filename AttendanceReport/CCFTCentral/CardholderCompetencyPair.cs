namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardholderCompetencyPair")]
    public partial class CardholderCompetencyPair
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CardholderID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int CompetencyID { get; set; }

        public Guid GlobalID { get; set; }

        public int isDisabled { get; set; }

        public DateTime? ExpiryDate { get; set; }

        public DateTime? EnableDate { get; set; }

        [StringLength(2500)]
        public string Comments { get; set; }

        public virtual Cardholder Cardholder { get; set; }

        public virtual Competency Competency { get; set; }
    }
}
