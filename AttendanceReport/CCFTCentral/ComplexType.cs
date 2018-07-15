namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ComplexType")]
    public partial class ComplexType
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ComplexType()
        {
            ComplexTypeProperties = new HashSet<ComplexTypeProperty>();
        }

        [Key]
        public Guid DataTypeID { get; set; }

        public virtual DataType DataType { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ComplexTypeProperty> ComplexTypeProperties { get; set; }
    }
}
