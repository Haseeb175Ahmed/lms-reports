using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport.EFERTDb
{
    public class EFERTDbContext : DbContext
    {
        public EFERTDbContext() : base("name=EFERTDb")
        {
            Database.SetInitializer<EFERTDbContext>(null);
        }

        public DbSet<CardHolderInfo> CardHolders { get; set; }

        public DbSet<CadreInfo> Cadres { get; set; }

        public DbSet<CrewInfo> Crews { get; set; }

        public DbSet<DepartmentInfo> Departments { get; set; }

        public DbSet<DesignationInfo> Designations { get; set; }

        public DbSet<SectionInfo> Sections { get; set; }

        public DbSet<CompanyInfo> Companies { get; set; }

        public DbSet<VisitorCardHolder> Visitors { get; set; }

        public DbSet<DailyCardHolder> DailyCardHolders { get; set; }

        public DbSet<BlockedPersonInfo> BlockedPersons { get; set; }

        public DbSet<CheckInAndOutInfo> CheckedInInfos { get; set; }

        public DbSet<EmailAddress> EmailAddresses { get; set; }

        public DbSet<VisitingLocations> VisitingLocations { get; set; }

        public DbSet<SystemSetting> SystemSetting { get; set; }

        public DbSet<AlertInfo> AlertInfos { get; set; }

        public DbSet<CategoryInfo> CategoryInfo { get; set; }

        public DbSet<Users> Users { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<VisitorCardHolder>().Property(p => p.Picture).HasColumnType("image");
            modelBuilder.Entity<DailyCardHolder>().Property(p => p.Picture).HasColumnType("image");
            modelBuilder.Entity<CardHolderInfo>().Property(p => p.Picture).HasColumnType("image");

            base.OnModelCreating(modelBuilder);
        }
    }
}
