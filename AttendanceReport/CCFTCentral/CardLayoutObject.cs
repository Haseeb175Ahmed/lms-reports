namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("CardLayoutObject")]
    public partial class CardLayoutObject
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public CardLayoutObject()
        {
            CardLayoutASCIIMifarePlus = new HashSet<CardLayoutASCIIMifarePlu>();
            CardLayoutObjectElements = new HashSet<CardLayoutObjectElement>();
        }

        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SideIndex { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ObjectIndex { get; set; }

        [Required]
        [StringLength(4)]
        public string ObjectType { get; set; }

        public int Left { get; set; }

        public int Top { get; set; }

        public int Right { get; set; }

        public int Bottom { get; set; }

        public int BorderStyle { get; set; }

        public int BorderColour { get; set; }

        public int FillColour { get; set; }

        public virtual CardLayoutBarcode CardLayoutBarcode { get; set; }

        public virtual CardLayoutGraphic CardLayoutGraphic { get; set; }

        public virtual CardLayoutMagstripe CardLayoutMagstripe { get; set; }

        public virtual CardLayoutMifare CardLayoutMifare { get; set; }

        public virtual CardLayoutMifarePlu CardLayoutMifarePlu { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardLayoutASCIIMifarePlu> CardLayoutASCIIMifarePlus { get; set; }

        public virtual CardLayoutSide CardLayoutSide { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardLayoutObjectElement> CardLayoutObjectElements { get; set; }

        public virtual CardLayoutSagem CardLayoutSagem { get; set; }

        public virtual CardLayoutSalto CardLayoutSalto { get; set; }

        public virtual CardLayoutText CardLayoutText { get; set; }
    }
}
