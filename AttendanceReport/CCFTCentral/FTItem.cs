namespace AttendanceReport.CCFTCentral
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("FTItem")]
    public partial class FTItem
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public FTItem()
        {
            AccessTriplets = new HashSet<AccessTriplet>();
            AccessZoneRelations = new HashSet<AccessZoneRelation>();
            AccessZoneRelations1 = new HashSet<AccessZoneRelation>();
            AlarmInstructionTags = new HashSet<AlarmInstructionTag>();
            AlarmZoneRelations = new HashSet<AlarmZoneRelation>();
            Cardholders = new HashSet<Cardholder>();
            CardholderRoles = new HashSet<CardholderRole>();
            ClubAlarmZones = new HashSet<ClubAlarmZone>();
            ConfigDispatchRoutes = new HashSet<ConfigDispatchRoute>();
            Divisions = new HashSet<Division>();
            Divisions1 = new HashSet<Division>();
            DivisionBadgeTexts = new HashSet<DivisionBadgeText>();
            DivisionFindVisitorGridColumns = new HashSet<DivisionFindVisitorGridColumn>();
            DivisionFindVisitorGridColumns1 = new HashSet<DivisionFindVisitorGridColumn>();
            DivisionHomePageFields = new HashSet<DivisionHomePageField>();
            DivisionHomePageFields1 = new HashSet<DivisionHomePageField>();
            DivisionMacroes = new HashSet<DivisionMacro>();
            DivisionMacroes1 = new HashSet<DivisionMacro>();
            DivisionVisitActions = new HashSet<DivisionVisitAction>();
            DivisionVisitorTypes = new HashSet<DivisionVisitorType>();
            DivisionVisitorTypes1 = new HashSet<DivisionVisitorType>();
            DivisionVisitorTypes2 = new HashSet<DivisionVisitorType>();
            DivisionVisitorTypeAccessGroups = new HashSet<DivisionVisitorTypeAccessGroup>();
            DivisionVisitorTypeAccessGroups1 = new HashSet<DivisionVisitorTypeAccessGroup>();
            DivisionVisitorTypeAccessGroups2 = new HashSet<DivisionVisitorTypeAccessGroup>();
            DivisionVisitorTypeFields = new HashSet<DivisionVisitorTypeField>();
            DivisionVisitorTypeFields1 = new HashSet<DivisionVisitorTypeField>();
            DivisionVisitorTypeFields2 = new HashSet<DivisionVisitorTypeField>();
            DivisionVisitorTypeHostTypes = new HashSet<DivisionVisitorTypeHostType>();
            DivisionVisitorTypeHostTypes1 = new HashSet<DivisionVisitorTypeHostType>();
            DivisionVisitorTypeHostTypes2 = new HashSet<DivisionVisitorTypeHostType>();
            DivisionVisitorTypeVisitVisitorFields = new HashSet<DivisionVisitorTypeVisitVisitorField>();
            DivisionVisitorTypeVisitVisitorFields1 = new HashSet<DivisionVisitorTypeVisitVisitorField>();
            DivisionVisitorTypeVisitVisitorFields2 = new HashSet<DivisionVisitorTypeVisitVisitorField>();
            DoorLastAccesses = new HashSet<DoorLastAccess>();
            DoorPersonalDataFieldPairs = new HashSet<DoorPersonalDataFieldPair>();
            DoorRelations = new HashSet<DoorRelation>();
            EventGroups = new HashSet<EventGroup>();
            ActionPlanTriggers = new HashSet<ActionPlanTrigger>();
            EDIConfigDatas = new HashSet<EDIConfigData>();
            FTItemRelationships = new HashSet<FTItemRelationship>();
            FTItemRelationships1 = new HashSet<FTItemRelationship>();
            GuardTours = new HashSet<GuardTour>();
            GuardTours1 = new HashSet<GuardTour>();
            GuardTourCheckpoints = new HashSet<GuardTourCheckpoint>();
            GuardTourProgresses = new HashSet<GuardTourProgress>();
            InterlockDoors = new HashSet<InterlockDoor>();
            InterlockExceptions = new HashSet<InterlockException>();
            InterlockExceptions1 = new HashSet<InterlockException>();
            LiftCarRelations = new HashSet<LiftCarRelation>();
            LogicBlockRelations = new HashSet<LogicBlockRelation>();
            MacroActions = new HashSet<MacroAction>();
            MacroSchedules = new HashSet<MacroSchedule>();
            MacroSchedules1 = new HashSet<MacroSchedule>();
            NotificationFilters = new HashSet<NotificationFilter>();
            NotificationFilters1 = new HashSet<NotificationFilter>();
            OperatorClubWorkstationPairs = new HashSet<OperatorClubWorkstationPair>();
            OperatorClubWorkstationPairs1 = new HashSet<OperatorClubWorkstationPair>();
            RemoteFTItemFields = new HashSet<RemoteFTItemField>();
            ReportScheduleTimes = new HashSet<ReportScheduleTime>();
            ReservedCardNumbers = new HashSet<ReservedCardNumber>();
            ReservedCardNumbers1 = new HashSet<ReservedCardNumber>();
            SaltoAccessTriplets = new HashSet<SaltoAccessTriplet>();
            SessionOperatorAccesses = new HashSet<SessionOperatorAccess>();
            SessionOperatorClubs = new HashSet<SessionOperatorClub>();
            SiteItems = new HashSet<SiteItem>();
            Visits = new HashSet<Visit>();
            Visits1 = new HashSet<Visit>();
            Visits2 = new HashSet<Visit>();
            VisitSegments = new HashSet<VisitSegment>();
            VisitSegments1 = new HashSet<VisitSegment>();
            VisitVisitors = new HashSet<VisitVisitor>();
            VisitVisitors1 = new HashSet<VisitVisitor>();
            VisitVisitors2 = new HashSet<VisitVisitor>();
            VisitVisitorAccessGroups = new HashSet<VisitVisitorAccessGroup>();
            VisitVisitorAccessGroups1 = new HashSet<VisitVisitorAccessGroup>();
            VisitVisitorAccessGroups2 = new HashSet<VisitVisitorAccessGroup>();
            VisitVisitorActions = new HashSet<VisitVisitorAction>();
            VisitVisitorActions1 = new HashSet<VisitVisitorAction>();
            Workstations = new HashSet<Workstation>();
            Workstations1 = new HashSet<Workstation>();
            AlarmKeypads = new HashSet<AlarmKeypad>();
            AlarmKeypads1 = new HashSet<AlarmKeypad>();
            Berts = new HashSet<Bert>();
            FTItem1 = new HashSet<FTItem>();
            FTItems = new HashSet<FTItem>();
            NotificationFilters2 = new HashSet<NotificationFilter>();
            FTItems1 = new HashSet<FTItem>();
            FTItem12 = new HashSet<FTItem>();
            FTItems2 = new HashSet<FTItem>();
            BulkChanges = new HashSet<BulkChange>();
            FTItem13 = new HashSet<FTItem>();
            FTItems3 = new HashSet<FTItem>();
        }

        public int ID { get; set; }

        [Required]
        [StringLength(110)]
        public string Name { get; set; }

        [Required]
        [StringLength(200)]
        public string Description { get; set; }

        public DateTime CreationTime { get; set; }

        public DateTime LastModificationTime { get; set; }

        public int LastModifyingOperatorID { get; set; }

        public DateTime? DeletionTime { get; set; }

        public int? DivisionID { get; set; }

        public int? TypeID { get; set; }

        [Required]
        [StringLength(16)]
        public string ShortName { get; set; }

        public int HomeServerID { get; set; }

        public Guid GlobalID { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AccessTriplet> AccessTriplets { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AccessZoneRelation> AccessZoneRelations { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AccessZoneRelation> AccessZoneRelations1 { get; set; }

        public virtual ActionPlan ActionPlan { get; set; }

        public virtual AlarmInstruction AlarmInstruction { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmInstructionTag> AlarmInstructionTags { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmZoneRelation> AlarmZoneRelations { get; set; }

        public virtual Archive Archive { get; set; }

        public virtual BulkChange BulkChange { get; set; }

        public virtual Calendar Calendar { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Cardholder> Cardholders { get; set; }

        public virtual Cardholder Cardholder { get; set; }

        public virtual CardholderLastSuccessfulAccess CardholderLastSuccessfulAccess { get; set; }

        public virtual CardholderLocation CardholderLocation { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CardholderRole> CardholderRoles { get; set; }

        public virtual CardType CardType { get; set; }

        public virtual Club Club { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ClubAlarmZone> ClubAlarmZones { get; set; }

        public virtual Competency Competency { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ConfigDispatchRoute> ConfigDispatchRoutes { get; set; }

        public virtual DayCategory DayCategory { get; set; }

        public virtual DeviceSoftware DeviceSoftware { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Division> Divisions { get; set; }

        public virtual Division Division { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Division> Divisions1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionBadgeText> DivisionBadgeTexts { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionFindVisitorGridColumn> DivisionFindVisitorGridColumns { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionFindVisitorGridColumn> DivisionFindVisitorGridColumns1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionHomePageField> DivisionHomePageFields { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionHomePageField> DivisionHomePageFields1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionMacro> DivisionMacroes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionMacro> DivisionMacroes1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitAction> DivisionVisitActions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorType> DivisionVisitorTypes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorType> DivisionVisitorTypes1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorType> DivisionVisitorTypes2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeAccessGroup> DivisionVisitorTypeAccessGroups { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeAccessGroup> DivisionVisitorTypeAccessGroups1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeAccessGroup> DivisionVisitorTypeAccessGroups2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeField> DivisionVisitorTypeFields { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeField> DivisionVisitorTypeFields1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeField> DivisionVisitorTypeFields2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeHostType> DivisionVisitorTypeHostTypes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeHostType> DivisionVisitorTypeHostTypes1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeHostType> DivisionVisitorTypeHostTypes2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeVisitVisitorField> DivisionVisitorTypeVisitVisitorFields { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeVisitVisitorField> DivisionVisitorTypeVisitVisitorFields1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DivisionVisitorTypeVisitVisitorField> DivisionVisitorTypeVisitVisitorFields2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DoorLastAccess> DoorLastAccesses { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DoorPersonalDataFieldPair> DoorPersonalDataFieldPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<DoorRelation> DoorRelations { get; set; }

        public virtual DVRCamera DVRCamera { get; set; }

        public virtual DVRType DVRType { get; set; }

        public virtual EDIConfiguration EDIConfiguration { get; set; }

        public virtual ElmoPersonalDataFieldPair ElmoPersonalDataFieldPair { get; set; }

        public virtual EnrolmentTemplate EnrolmentTemplate { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EventGroup> EventGroups { get; set; }

        public virtual FingerprintBiometricType FingerprintBiometricType { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ActionPlanTrigger> ActionPlanTriggers { get; set; }

        public virtual BertDependent BertDependent { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EDIConfigData> EDIConfigDatas { get; set; }

        public virtual FTItemType FTItemType { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItemRelationship> FTItemRelationships { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItemRelationship> FTItemRelationships1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<GuardTour> GuardTours { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<GuardTour> GuardTours1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<GuardTourCheckpoint> GuardTourCheckpoints { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<GuardTourProgress> GuardTourProgresses { get; set; }

        public virtual IconSet IconSet { get; set; }

        public virtual InsertionTag InsertionTag { get; set; }

        public virtual IntercomSystem IntercomSystem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<InterlockDoor> InterlockDoors { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<InterlockException> InterlockExceptions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<InterlockException> InterlockExceptions1 { get; set; }

        public virtual ISIntercom ISIntercom { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<LiftCarRelation> LiftCarRelations { get; set; }

        public virtual LogicBlock LogicBlock { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<LogicBlockRelation> LogicBlockRelations { get; set; }

        public virtual Macro Macro { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MacroAction> MacroActions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MacroSchedule> MacroSchedules { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<MacroSchedule> MacroSchedules1 { get; set; }

        public virtual Note Note { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<NotificationFilter> NotificationFilters { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<NotificationFilter> NotificationFilters1 { get; set; }

        public virtual Operator Operator { get; set; }

        public virtual OperatorClub OperatorClub { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubWorkstationPair> OperatorClubWorkstationPairs { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<OperatorClubWorkstationPair> OperatorClubWorkstationPairs1 { get; set; }

        public virtual OperatorPrivilege OperatorPrivilege { get; set; }

        public virtual Panel Panel { get; set; }

        public virtual PersonalDataField PersonalDataField { get; set; }

        public virtual Picture Picture { get; set; }

        public virtual Reception Reception { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RemoteFTItemField> RemoteFTItemFields { get; set; }

        public virtual ReportDescription ReportDescription { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ReportScheduleTime> ReportScheduleTimes { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ReservedCardNumber> ReservedCardNumbers { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ReservedCardNumber> ReservedCardNumbers1 { get; set; }

        public virtual Role Role { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SaltoAccessTriplet> SaltoAccessTriplets { get; set; }

        public virtual Schedule Schedule { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SessionOperatorAccess> SessionOperatorAccesses { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SessionOperatorClub> SessionOperatorClubs { get; set; }

        public virtual SiteItem SiteItem { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SiteItem> SiteItems { get; set; }

        public virtual SitePlan SitePlan { get; set; }

        public virtual Sound Sound { get; set; }

        public virtual UniversalCardFormat UniversalCardFormat { get; set; }

        public virtual Viewer Viewer { get; set; }

        public virtual Visit Visit { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Visit> Visits { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Visit> Visits1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Visit> Visits2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitSegment> VisitSegments { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitSegment> VisitSegments1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitor> VisitVisitors { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitor> VisitVisitors1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitor> VisitVisitors2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitorAccessGroup> VisitVisitorAccessGroups { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitorAccessGroup> VisitVisitorAccessGroups1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitorAccessGroup> VisitVisitorAccessGroups2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitorAction> VisitVisitorActions { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<VisitVisitorAction> VisitVisitorActions1 { get; set; }

        public virtual WizardTemplate WizardTemplate { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Workstation> Workstations { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Workstation> Workstations1 { get; set; }

        public virtual Workstation Workstation { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmKeypad> AlarmKeypads { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<AlarmKeypad> AlarmKeypads1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Bert> Berts { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItem1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItems { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<NotificationFilter> NotificationFilters2 { get; set; }

        public virtual KeyMap KeyMap { get; set; }

        public virtual FTItem FTItem11 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItems1 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItem12 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItems2 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<BulkChange> BulkChanges { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItem13 { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<FTItem> FTItems3 { get; set; }
    }
}
