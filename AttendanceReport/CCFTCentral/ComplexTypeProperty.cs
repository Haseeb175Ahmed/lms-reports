namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ComplexTypeProperty")]
    public partial class ComplexTypeProperty
    {
        [Key]
        [Column(Order = 0)]
        public Guid DataTypeID { get; set; }

        [Key]
        [Column(Order = 1)]
        public Guid PropertyID { get; set; }

        public bool Replicated { get; set; }

        public virtual ComplexType ComplexType { get; set; }

        public virtual Property Property { get; set; }
    }
}
