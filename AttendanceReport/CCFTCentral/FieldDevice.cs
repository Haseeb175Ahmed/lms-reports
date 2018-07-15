namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("FieldDevice")]
    public partial class FieldDevice
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public FieldDevice()
        {
            ActionPlanOutputs = new HashSet<ActionPlanOutput>();
            FieldDeviceSoftwares = new HashSet<FieldDeviceSoftware>();
            ReflectIOPairs = new HashSet<ReflectIOPair>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int? ParentDeviceID { get; set; }

        public int PhysicalAddress { get; set; }

        public virtual AperioDoor AperioDoor { get; set; }

        public virtual DisturbanceSensor DisturbanceSensor { get; set; }

        public virtual ExternalEventSource ExternalEventSource { get; set; }

        public virtual ExternalSystem ExternalSystem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ActionPlanOutput> ActionPlanOutputs { get; set; }

        public virtual SiteItem SiteItem { get; set; }

        public virtual SiteItem SiteItem1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FieldDeviceSoftware> FieldDeviceSoftwares { get; set; }

        public virtual Input Input { get; set; }

        public virtual LiftHLI LiftHLI { get; set; }

        public virtual Output Output { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ReflectIOPair> ReflectIOPairs { get; set; }

        public virtual TautWireSensor TautWireSensor { get; set; }
    }
}
