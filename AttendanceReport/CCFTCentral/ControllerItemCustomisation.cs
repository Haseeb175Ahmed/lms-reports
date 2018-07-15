namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ControllerItemCustomisation")]
    public partial class ControllerItemCustomisation
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ControllerItemTypeId { get; set; }

        [Key]
        [Column(Order = 1)]
        public Guid ComponentId { get; set; }
    }
}
