namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ExternalSystem")]
    public partial class ExternalSystem
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public ExternalSystem()
        {
            ExternalSystem1 = new HashSet<ExternalSystem>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [StringLength(255)]
        public string HostName { get; set; }

        [Required]
        [StringLength(16)]
        public string IPAddress { get; set; }

        public int? IdentitySameAs { get; set; }

        [Required]
        [StringLength(64)]
        public string SystemID { get; set; }

        [Required]
        [StringLength(64)]
        public string SystemType { get; set; }

        [Required]
        [StringLength(1024)]
        public string SystemConfig { get; set; }

        public virtual FieldDevice FieldDevice { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ExternalSystem> ExternalSystem1 { get; set; }

        public virtual ExternalSystem ExternalSystem2 { get; set; }
    }
}
