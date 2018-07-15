namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DataType")]
    public partial class DataType
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public DataType()
        {
            Properties = new HashSet<Property>();
        }

        public Guid DataTypeID { get; set; }

        public int BuiltInTypeID { get; set; }

        public int NameID { get; set; }

        [Required]
        [StringLength(50)]
        public string Name { get; set; }

        public virtual ComplexType ComplexType { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Property> Properties { get; set; }
    }
}
