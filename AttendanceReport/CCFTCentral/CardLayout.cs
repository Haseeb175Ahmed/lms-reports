namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardLayout")]
    public partial class CardLayout
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public CardLayout()
        {
            CardLayoutSides = new HashSet<CardLayoutSide>();
        }

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Column(TypeName = "image")]
        public byte[] layoutData { get; set; }

        public int Left { get; set; }

        public int Top { get; set; }

        public int Right { get; set; }

        public int Bottom { get; set; }

        public virtual CardType CardType { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardLayoutSide> CardLayoutSides { get; set; }
    }
}
