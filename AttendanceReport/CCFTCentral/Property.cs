namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Property")]
    public partial class Property
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Property()
        {
            ComplexTypeProperties = new HashSet<ComplexTypeProperty>();
        }

        public Guid PropertyID { get; set; }

        public int FieldID { get; set; }

        public Guid DataTypeID { get; set; }

        public bool Replicated { get; set; }

        public int NameID { get; set; }

        [Required]
        [StringLength(110)]
        public string Name { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ComplexTypeProperty> ComplexTypeProperties { get; set; }

        public virtual DataType DataType { get; set; }
    }
}
