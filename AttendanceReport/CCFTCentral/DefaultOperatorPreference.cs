namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DefaultOperatorPreference")]
    public partial class DefaultOperatorPreference
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int TypeCode { get; set; }

        [Required]
        [MaxLength(1000)]
        public byte[] StateData { get; set; }
    }
}
