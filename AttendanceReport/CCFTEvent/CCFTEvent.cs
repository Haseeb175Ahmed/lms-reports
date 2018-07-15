namespace AttendanceReport.CCFTEvent
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class CCFTEvent : DbContext
    {
        public CCFTEvent()
            : base("name=CCFTEvent")
        {
            Database.SetInitializer<CCFTEvent>(null);
        }

        public virtual DbSet<AlarmStack> AlarmStacks { get; set; }
        public virtual DbSet<AlarmUpdate> AlarmUpdates { get; set; }
        public virtual DbSet<CardEvent> CardEvents { get; set; }
        public virtual DbSet<DVRData> DVRDatas { get; set; }
        public virtual DbSet<Event> Events { get; set; }
        public virtual DbSet<EventReturnToNormal> EventReturnToNormals { get; set; }
        public virtual DbSet<Image> Images { get; set; }
        public virtual DbSet<Lookup> Lookups { get; set; }
        public virtual DbSet<OnlineArchiveFile> OnlineArchiveFiles { get; set; }
        public virtual DbSet<SALTOLastEvent> SALTOLastEvents { get; set; }
        public virtual DbSet<Sequence> Sequences { get; set; }
        public virtual DbSet<Trigger> Triggers { get; set; }
        public virtual DbSet<AlarmHistory> AlarmHistories { get; set; }
        public virtual DbSet<NotificationAlarm> NotificationAlarms { get; set; }
        public virtual DbSet<RelatedItem> RelatedItems { get; set; }
        public virtual DbSet<SingleValue> SingleValues { get; set; }
        public virtual DbSet<TriggerSequenceLocation> TriggerSequenceLocations { get; set; }
        public virtual DbSet<TriggerSequencePair> TriggerSequencePairs { get; set; }
        public virtual DbSet<UnlinkedAck> UnlinkedAcks { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Event>()
                .HasOptional(e => e.AlarmStack)
                .WithRequired(e => e.Event);

            modelBuilder.Entity<Event>()
                .HasOptional(e => e.AlarmUpdate)
                .WithRequired(e => e.Event)
                .WillCascadeOnDelete();

            modelBuilder.Entity<Event>()
                .HasOptional(e => e.CardEvent)
                .WithRequired(e => e.Event)
                .WillCascadeOnDelete();

            modelBuilder.Entity<Sequence>()
                .HasOptional(e => e.DVRData)
                .WithRequired(e => e.Sequence)
                .WillCascadeOnDelete();

            modelBuilder.Entity<Sequence>()
                .HasMany(e => e.Images)
                .WithOptional(e => e.Sequence)
                .WillCascadeOnDelete();
        }
    }
}
