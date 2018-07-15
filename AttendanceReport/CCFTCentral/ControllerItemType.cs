namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ControllerItemType")]
    public partial class ControllerItemType
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ControllerItemType()
        {
            ConfigDispatchFieldControllerItemTypePairs = new HashSet<ConfigDispatchFieldControllerItemTypePair>();
            ConfigDispatchRoutes = new HashSet<ConfigDispatchRoute>();
        }

        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ID { get; set; }

        public int BertDatabaseID { get; set; }

        public Guid ClassID { get; set; }

        public int Flags { get; set; }

        [Required]
        [StringLength(50)]
        public string Name { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ConfigDispatchFieldControllerItemTypePair> ConfigDispatchFieldControllerItemTypePairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ConfigDispatchRoute> ConfigDispatchRoutes { get; set; }
    }
}
