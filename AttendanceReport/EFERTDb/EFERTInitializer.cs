using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity.Utilities;
using System.IO;
using System.Xml;
using System.Data.Entity.Infrastructure.DependencyResolution;
using System.Data.Entity.Infrastructure;

namespace AttendanceReport.EFERTDb
{
    public class EFERTInitializer : CreateDatabaseIfNotExists<EFERTDbContext>
    {
        protected override void Seed(EFERTDbContext context)
        {
            #region Plant Locations

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "SECURITY BUILDING",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MARKETING",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "ADMIN BUILDING",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "WORKSHOP",
                IsOnPlant = true
            });
            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "SAP BUILDING",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "WAREHOUSE",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "TECHNICAL BUILDING",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "LABORATORY",
                IsOnPlant = true
            });
            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "ENVEN BUILDING",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "PM&S-1",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "PM&S-II",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "PM&S-III",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "UREA-1",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "UREA-II",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "AMMONIA-II",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "CCR",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "UREA-II MAINTENANCE",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "CLIENT WAREHOUSE",
                IsOnPlant = true
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "TRAINING BUILDING",
                IsOnPlant = true
            });

            #endregion

            #region Colony Location

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "NEW MARKET",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "OLD MARKET",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT CLUB",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "GRAMMAR SCHOOL",
                IsOnPlant = false
            });
            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT LAUNDRY",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT CRICKET STADIUM",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT SCHOOL GROUND",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT GUEST HOUSE",
                IsOnPlant = false
            });
            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT BOQ-1",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT BOQ-II",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT BOQ-III",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT BOQ-IV",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MANAGEMENT MARIT SWEET",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "ENGRO COLLAGE",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "MASJID",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "EMPLOYEES LAUNDRY",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "EMPLOYEES CLUB",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "OLD EMPLOYEES STADIUM",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "NEW EMPLOYEES STADIUM",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "EMPLOYEES COMMUNITY GROUND",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "EMPLOYEES ETP",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "EMPLOYEES GUEST HOUSE",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "EMPLOYEES BOQ-1",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "EMPLOYEES BOQ-II",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "EMPLOYEES X-BUILDING",
                IsOnPlant = false
            });

            context.VisitingLocations.Add(new VisitingLocations()
            {
                Location = "EMPLOYEES DOUBLE STORY BOQ",
                IsOnPlant = false
            });

            #endregion

            #region Departments

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "ADMINISTRATION"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "HSE&T"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "I&E"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "MAINTENANCE"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "MARKETING"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "MECHANICAL ENGINGEERING"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "MEDICAL"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "PLANT MANAGEMENT"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "PROCESS ENGINEERING"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "PRODUCTION"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "P&MM"
            });

            context.Departments.Add(new DepartmentInfo()
            {
                DepartmentName = "OTHERS"
            });

            #endregion

            #region Sections

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "ADMINISTRATION"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "ACCOUNTS"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "AUDITS"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "HOUSING"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "P&IR"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "SECURITY"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "MECHANICAL ENGINEERING"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "INSPECTION"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "MTS"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "PROJECTS"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "HSE&T"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "ENVIRONMENT"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "SAFETY"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "MARKETING"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "DISTRIBUTION"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "SALES"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "PROCESS ENGINEERING"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "LABORATORY"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "PROCESS"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "I&E"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "ELECTRICAL"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "INSTRUMENT"
            });

            context.Sections.Add(new SectionInfo()
            {
                SectionName = "SYSTEM"
            });

            #endregion

            #region Company

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "A KAREEM & SONS"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "A HAMEED & COMPANY"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "ABDUL GHAFOOR & SONS"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "ABDUL MAJEED & SONS"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "AFZAL & COMPANY"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "AHMED & COMPANY"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "AL FUKRA TRADERS"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "AL AZAAN ENGINEERING"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "AL MUNIR FABRICATION"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "AL MUSTAFA & COMPANY"
            });

            context.Companies.Add(new CompanyInfo()
            {
                CompanyName = "AL RAZA FURNITURE"
            });

            #endregion

            #region Designation

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "TWM"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "USL"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "SWEEPER"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "WAITER"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "DRIVER"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "COMPUTER OPERATER"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "MESSENGER"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "TECHNICIAN"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "ELECTRICAL TECHNICIAN"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "INSTRUMENT TECHNICIAN"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "OFFICE  ASSISTANT"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "GARDNER"
            });

            context.Designations.Add(new DesignationInfo()
            {
                Designation = "AC TECHNICIAN"
            });

            #endregion

            #region Cadre

            context.Cadres.Add(new CadreInfo()
            {
                CadreName = "MPT"
            });

            context.Cadres.Add(new CadreInfo()
            {
                CadreName = "NMPT"
            });

            context.Cadres.Add(new CadreInfo()
            {
                CadreName = "GTE"
            });

            context.Cadres.Add(new CadreInfo()
            {
                CadreName = "TAP"
            });

            context.Cadres.Add(new CadreInfo()
            {
                CadreName = "WORK ORDER"
            });

            context.Cadres.Add(new CadreInfo()
            {
                CadreName = "EC"
            });

            context.Cadres.Add(new CadreInfo()
            {
                CadreName = "CONTRACTOR"
            });

            context.Cadres.Add(new CadreInfo()
            {
                CadreName = "SUPERVISOR"
            });
            
            #endregion

            base.Seed(context);
        }
    }
}
