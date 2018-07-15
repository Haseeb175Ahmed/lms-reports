namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Tile")]
    public partial class Tile
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Tile()
        {
            TileItemRelationships = new HashSet<TileItemRelationship>();
        }

        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        public Guid GlobalID { get; set; }

        public Guid ClassID { get; set; }

        [Required]
        [StringLength(200)]
        public string ConfigContractName { get; set; }

        public double Left { get; set; }

        public double Top { get; set; }

        public double Width { get; set; }

        public double Height { get; set; }

        [Required]
        [StringLength(200)]
        public string ControlPresenterContractName { get; set; }

        [Required]
        [StringLength(200)]
        public string ControlConfigPresenterContractName { get; set; }

        public virtual Panel Panel { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<TileItemRelationship> TileItemRelationships { get; set; }
    }
}
