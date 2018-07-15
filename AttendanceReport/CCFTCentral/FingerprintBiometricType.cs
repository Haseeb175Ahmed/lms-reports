namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("FingerprintBiometricType")]
    public partial class FingerprintBiometricType
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int FTItemID { get; set; }

        public int FacilityCode { get; set; }

        public int FirstFingerID { get; set; }

        public int SecondFingerID { get; set; }

        public int FirstDuressFingerID { get; set; }

        public int SecondDuressFingerID { get; set; }

        public bool SaveFingerprints { get; set; }

        public bool AllowOperatorOverride { get; set; }

        public int StartingSector { get; set; }

        [Required]
        [StringLength(12)]
        public string MifareKeyA { get; set; }

        public bool FingerprintMatching { get; set; }

        public int MatchingThreshold { get; set; }

        public virtual FTItem FTItem { get; set; }
    }
}
