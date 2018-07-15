namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("LiftCar")]
    public partial class LiftCar
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public LiftCar()
        {
            LiftCarLevels = new HashSet<LiftCarLevel>();
            LiftCarRelations = new HashSet<LiftCarRelation>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int FloorSelectTimeout { get; set; }

        public int PinEntryTimeout { get; set; }

        public int CardSwipeTimeout { get; set; }

        public bool FailFree { get; set; }

        public int GroupNumber { get; set; }

        public int CarNumber { get; set; }

        public bool DisableFeedback { get; set; }

        public virtual SiteItem SiteItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<LiftCarLevel> LiftCarLevels { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<LiftCarRelation> LiftCarRelations { get; set; }
    }
}
