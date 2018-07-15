namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("DeviceSoftware")]
    public partial class DeviceSoftware
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public DeviceSoftware()
        {
            DeviceSoftwareModules = new HashSet<DeviceSoftwareModule>();
            FieldDeviceSoftwares = new HashSet<FieldDeviceSoftware>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Required]
        [StringLength(12)]
        public string Version { get; set; }

        [Required]
        [StringLength(12)]
        public string MinimumControllerVersion { get; set; }

        [Required]
        [StringLength(12)]
        public string MinimumHeadEndVersion { get; set; }

        [Column(TypeName = "image")]
        public byte[] FileData { get; set; }

        public virtual FTItem FTItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DeviceSoftwareModule> DeviceSoftwareModules { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FieldDeviceSoftware> FieldDeviceSoftwares { get; set; }
    }
}
