namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class SingleValue
    {
        [Key]
        [Column(Order = 0)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Version { get; set; }

        [Key]
        [Column(Order = 1)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SystemOperatorID { get; set; }

        [Key]
        [Column(Order = 2)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int RootDivisionID { get; set; }

        [StringLength(260)]
        public string ArchiveLocation { get; set; }

        public int? SiteSerialNumber { get; set; }

        [Key]
        [Column(Order = 3)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ChallengeShowFirstName { get; set; }

        [Key]
        [Column(Order = 4)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ChallengeShowLastName { get; set; }

        [Key]
        [Column(Order = 5)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ChallengeShowDescription { get; set; }

        [Key]
        [Column(Order = 6)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ChallengeShowCardNumber { get; set; }

        [Key]
        [Column(Order = 7)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ChallengeOperatorComment { get; set; }

        [Key]
        [Column(Order = 8)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int ChallengeAudioType { get; set; }

        [Key]
        [Column(Order = 9)]
        [StringLength(260)]
        public string ChallengeAudioFile { get; set; }

        [Column(TypeName = "image")]
        public byte[] ChallengeAudioData { get; set; }

        [Key]
        [Column(Order = 10)]
        [StringLength(5)]
        public string SiteACCT { get; set; }

        [Key]
        [Column(Order = 11)]
        [MaxLength(32)]
        public byte[] Validation { get; set; }

        [Key]
        [Column(Order = 12)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int StateMapping { get; set; }

        [Key]
        [Column(Order = 13)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AlarmNotesPriority { get; set; }

        [StringLength(32)]
        public string MifareSiteKey { get; set; }

        [Key]
        [Column(Order = 14)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int AllowUSBUpgrades { get; set; }

        [Key]
        [Column(Order = 15)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int DownloadLimit { get; set; }

        [Key]
        [Column(Order = 16)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int VisitPurgeTimeOfDay { get; set; }

        [Key]
        [Column(Order = 17)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int VisitPurgeAgeInDays { get; set; }

        [Key]
        [Column(Order = 18)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyEmailOutEnabled { get; set; }

        [Key]
        [Column(Order = 19)]
        [StringLength(50)]
        public string NotifyEmailOutAccount { get; set; }

        [Key]
        [Column(Order = 20)]
        [StringLength(50)]
        public string NotifyEmailOutServer { get; set; }

        [Key]
        [Column(Order = 21)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyEmailOutPort { get; set; }

        [Key]
        [Column(Order = 22)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyEmailOutSSL { get; set; }

        [Key]
        [Column(Order = 23)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifySMSOutEnabled { get; set; }

        [Key]
        [Column(Order = 24)]
        [StringLength(1024)]
        public string NotifySMSOutDomain { get; set; }

        [Key]
        [Column(Order = 25)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int SMSRequiresEmail { get; set; }

        [Key]
        [Column(Order = 26)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifySMSOutFormat { get; set; }

        [Key]
        [Column(Order = 27)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyExpiryTime { get; set; }

        [Key]
        [Column(Order = 28)]
        public DateTime LastExpiryCheck { get; set; }

        [Key]
        [Column(Order = 29)]
        public bool EnableConfirmedAlarms { get; set; }

        [Key]
        [Column(Order = 30)]
        public bool DisableAlarmIndications { get; set; }

        [Key]
        [Column(Order = 31)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyEmailInEnabled { get; set; }

        [Key]
        [Column(Order = 32)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyEmailInProtocol { get; set; }

        [Key]
        [Column(Order = 33)]
        [StringLength(50)]
        public string NotifyEmailInAccount { get; set; }

        [Key]
        [Column(Order = 34)]
        [StringLength(50)]
        public string NotifyEmailInServer { get; set; }

        [Key]
        [Column(Order = 35)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyEmailInPort { get; set; }

        [Key]
        [Column(Order = 36)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyEmailInSSL { get; set; }

        [Key]
        [Column(Order = 37)]
        [StringLength(32)]
        public string MifarePlusKey { get; set; }

        [Key]
        [Column(Order = 38)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int UsesMifarePlusSL3 { get; set; }

        [Key]
        [Column(Order = 39)]
        [MaxLength(200)]
        public byte[] NotifyEmailOutPassword { get; set; }

        [Key]
        [Column(Order = 40)]
        [MaxLength(200)]
        public byte[] NotifyEmailInPassword { get; set; }

        [Key]
        [Column(Order = 41)]
        public bool DelayDialling { get; set; }

        [Key]
        [Column(Order = 42)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyEmailInInterval { get; set; }

        [Key]
        [Column(Order = 43)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifyForwardAcks { get; set; }

        [Key]
        [Column(Order = 44)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int EnableKeyLoading { get; set; }

        [Key]
        [Column(Order = 45)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int UseDefaultKey { get; set; }

        [Key]
        [Column(Order = 46)]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int NotifySMSOutItemNameFormat { get; set; }

        [Key]
        [Column(Order = 47)]
        [StringLength(100)]
        public string NotifyEmailOutSender { get; set; }
    }
}
