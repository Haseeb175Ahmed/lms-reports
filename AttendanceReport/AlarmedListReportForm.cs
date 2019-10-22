using AttendanceReport.EFERTDb;
using iText.Kernel.Colors;
using iText.Kernel.Events;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceReport
{
    public partial class AlarmedListReportForm : Form
    {
        private Dictionary<string, List<CardHolderReportInfo>> mLstCardHolders = null;

        public AlarmedListReportForm()
        {
            InitializeComponent();
            EFERTDbUtility.InitializeDatabases();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            this.mLstCardHolders = new Dictionary<string, List<CardHolderReportInfo>>();

            var lstCheckedInChls = (from checkIn in EFERTDbUtility.mEFERTDb.CheckedInInfos
                                    where checkIn != null 
                                    select new {
                                        CardHolderId = checkIn.CardHolderId,
                                        DailyCardHolderId = checkIn.DailyCardHolderId,
                                        VisitorId = checkIn.VisitorId
                                    }).Distinct().ToList();

            List<int> cardHolderIds = (from checkedInChl in lstCheckedInChls
                                       where checkedInChl != null && checkedInChl.CardHolderId != null
                                       select Convert.ToInt32(checkedInChl.CardHolderId)).ToList();

            List<int> dailyCardHolderIds = (from checkedInChl in lstCheckedInChls
                                           where checkedInChl != null && checkedInChl.DailyCardHolderId != null
                                           select Convert.ToInt32(checkedInChl.DailyCardHolderId)).ToList();

            List<int> visitorIds = (from checkedInChl in lstCheckedInChls
                                            where checkedInChl != null && checkedInChl.DailyCardHolderId != null
                                            select Convert.ToInt32(checkedInChl.DailyCardHolderId)).ToList();

            List<EFERTDb.CardHolderInfo> chls = (from chl in EFERTDbUtility.mEFERTDb.CardHolders
                                                 where chl != null && cardHolderIds.Contains(chl.CardHolderId)
                                                 select chl).ToList();

            List<DailyCardHolder> dailyChls = (from dailychl in EFERTDbUtility.mEFERTDb.DailyCardHolders
                                                 where dailychl != null && dailyCardHolderIds.Contains(dailychl.DailyCardHolderId)
                                                 select dailychl).ToList();

            List<VisitorCardHolder> visitorChls = (from visitor in EFERTDbUtility.mEFERTDb.Visitors
                                                 where visitor != null && visitorIds.Contains(visitor.VisitorId)
                                                 select visitor).ToList();


            foreach (EFERTDb.CardHolderInfo cardHolder in chls)
            {
                int daysCount = 0;

                LimitStatus limitStatus = EFERTDbUtility.CheckIfUserCheckedInLimitReached(cardHolder.CheckInInfos, cardHolder.BlockingInfos, ref daysCount, false);

                if (limitStatus == LimitStatus.LimitReached || limitStatus == LimitStatus.EmailAlerted)
                {
                    string category = cardHolder.ConstractorInfo;
                    category = string.IsNullOrEmpty(category) ? "Unknown" : category;
                    string firstName = cardHolder.FirstName;
                    string cnicNumber = cardHolder.CNICNumber;
                    string companyName = !Object.ReferenceEquals(null, cardHolder.Company) ?  cardHolder.Company.CompanyName : "Test"; 
                    string blockedStatus = limitStatus == LimitStatus.LimitReached? "Blocked" : "Alarmed" ;


                    if (this.mLstCardHolders.ContainsKey(category))
                    {
                        this.mLstCardHolders[category].Add(new CardHolderReportInfo()
                        {
                            FirstName = firstName,
                            CNICNumber = cnicNumber,
                            BlockedStatus = blockedStatus,
                            Category = category,
                            Company = companyName,
                            TotalDaysCkekin = daysCount
                        });

                    }
                    else
                    {
                        List<CardHolderReportInfo> lstChls = new List<CardHolderReportInfo>();
                        lstChls.Add(new CardHolderReportInfo()
                        {
                            FirstName = firstName,
                            CNICNumber = cnicNumber,
                            BlockedStatus = blockedStatus,
                            Category = category,
                            Company = companyName,
                            TotalDaysCkekin = daysCount
                        });


                        this.mLstCardHolders.Add(category, lstChls);
                    }
                }
            }

            foreach (EFERTDb.DailyCardHolder dailyCardHolder in dailyChls)
            {
                int daysCount = 0;

                LimitStatus limitStatus = EFERTDbUtility.CheckIfUserCheckedInLimitReached(dailyCardHolder.CheckInInfos, dailyCardHolder.BlockingInfos, ref daysCount, false);

                if (limitStatus == LimitStatus.LimitReached || limitStatus == LimitStatus.EmailAlerted)
                {
                    string category = dailyCardHolder.ConstractorInfo;
                    category = string.IsNullOrEmpty(category) ? "Unknown" : category;
                    string firstName = dailyCardHolder.FirstName;
                    string cnicNumber = dailyCardHolder.CNICNumber;
                    string companyName = dailyCardHolder.CompanyName;
                    string blockedStatus = limitStatus == LimitStatus.LimitReached ? "Blocked" : "Alarmed";


                    if (this.mLstCardHolders.ContainsKey(category))
                    {
                        this.mLstCardHolders[category].Add(new CardHolderReportInfo()
                        {
                            FirstName = firstName,
                            CNICNumber = cnicNumber,
                            BlockedStatus = blockedStatus,
                            Category = category,
                            Company = companyName,
                            TotalDaysCkekin = daysCount
                        });

                    }
                    else
                    {
                        List<CardHolderReportInfo> lstChls = new List<CardHolderReportInfo>();
                        lstChls.Add(new CardHolderReportInfo()
                        {
                            FirstName = firstName,
                            CNICNumber = cnicNumber,
                            BlockedStatus = blockedStatus,
                            Category = category,
                            Company = companyName,
                            TotalDaysCkekin = daysCount
                        });


                        this.mLstCardHolders.Add(category, lstChls);
                    }
                }
            }

            foreach (EFERTDb.VisitorCardHolder visitor in visitorChls)
            {
                int daysCount = 0;

                LimitStatus limitStatus = EFERTDbUtility.CheckIfUserCheckedInLimitReached(visitor.CheckInInfos, visitor.BlockingInfos, ref daysCount, false);

                if (limitStatus == LimitStatus.LimitReached || limitStatus == LimitStatus.EmailAlerted)
                {
                    string category = visitor.VisitorInfo;
                    category = string.IsNullOrEmpty(category) ? "Unknown" : category;
                    string firstName = visitor.FirstName;
                    string cnicNumber = visitor.CNICNumber;
                    string companyName = visitor.CompanyName;
                    string blockedStatus = limitStatus == LimitStatus.LimitReached ? "Blocked" : "Alarmed";


                    if (this.mLstCardHolders.ContainsKey(category))
                    {
                        this.mLstCardHolders[category].Add(new CardHolderReportInfo()
                        {
                            FirstName = firstName,
                            CNICNumber = cnicNumber,
                            BlockedStatus = blockedStatus,
                            Category = category,
                            Company = companyName,
                            TotalDaysCkekin= daysCount
                        });

                    }
                    else
                    {
                        List<CardHolderReportInfo> lstChls = new List<CardHolderReportInfo>();
                        lstChls.Add(new CardHolderReportInfo()
                        {
                            FirstName = firstName,
                            CNICNumber = cnicNumber,
                            BlockedStatus = blockedStatus,
                            Category = category,
                            Company = companyName,
                            TotalDaysCkekin = daysCount
                        });


                        this.mLstCardHolders.Add(category, lstChls);
                    }
                }
            }

            if (this.mLstCardHolders != null && this.mLstCardHolders.Count > 0)
            {
                this.saveFileDialog1.ShowDialog(this);
            }
            else
            {

                MessageBox.Show(this, "No data exist on current selected date range.");
            }

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

            string extension = Path.GetExtension(this.saveFileDialog1.FileName);

            if (extension == ".pdf")
            {
                this.SaveAsPdf(this.mLstCardHolders, "Alarmed List Report");
            }
            else if (extension == ".xlsx")
            {
                this.SaveAsExcel(this.mLstCardHolders, "Alarmed List Report", "Alarmed List Report");
            }
        }

        private void SaveAsPdf(Dictionary<string, List<CardHolderReportInfo>> data, string heading)
        {
            Cursor currentCursor = Cursor.Current;

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (data != null)
                {
                    using (PdfWriter pdfWriter = new PdfWriter(this.saveFileDialog1.FileName))
                    {
                        using (PdfDocument pdfDocument = new PdfDocument(pdfWriter))
                        {
                            using (Document doc = new Document(pdfDocument))
                            {
                                doc.SetFont(PdfFontFactory.CreateFont("Fonts/SEGOEUIL.TTF"));
                                string headerLeftText = "Alarmed List Report";
                                string headerRightText = string.Empty;
                                string footerLeftText = "This is computer generated report.";
                                string footerRightText = "Report generated on: " + DateTime.Now.ToString();

                                pdfDocument.AddEventHandler(PdfDocumentEvent.START_PAGE, new PdfHeaderAndFooter(doc, true, headerLeftText, headerRightText));
                                pdfDocument.AddEventHandler(PdfDocumentEvent.END_PAGE, new PdfHeaderAndFooter(doc, false, footerLeftText, footerRightText));

                                //pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(1000F, 1000F));
                                //Table table = new Table((new List<float>() { 100F, 100F, 100F, 200F, 70F,70F,100F, 100F, 100F, 70F, 200F, 100F, 70F, 60F }).ToArray());
                                pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(1080F, 842F));
                                Table table = new Table((new List<float>() { 90F, 140F, 250F, 120F, 240F, 120F }).ToArray());

                                table.SetWidth(980F);
                                table.SetFixedLayout();

                                //Table table = new Table((new List<float>() { 8F, 100F, 150F, 225F, 60F, 40F, 100F, 125F, 150F }).ToArray());

                                this.AddMainHeading(table, heading);

                                this.AddNewEmptyRow(table);
                                //this.AddNewEmptyRow(table);

                                //Sections and Data


                                foreach (KeyValuePair<string, List<CardHolderReportInfo>> category in data)
                                {
                                    if (category.Value == null)
                                    {
                                        continue;
                                    }

                                    //Cadre
                                    this.AddCategoryRow(table, category.Key);

                                    //Data
                                    //this.AddNewEmptyRow(table, false);

                                    this.AddTableHeaderRow(table);

                                    for (int i = 0; i < category.Value.Count; i++)
                                    {
                                        CardHolderReportInfo chl = category.Value[i];
                                        this.AddTableDataRow(table, chl, i % 2 == 0);
                                    }

                                    this.AddNewEmptyRow(table);
                                }



                                doc.Add(table);

                                doc.Close();
                            }
                        }

                        System.Diagnostics.Process.Start(this.saveFileDialog1.FileName);
                    }
                }
                Cursor.Current = currentCursor;
            }
            catch (Exception exp)
            {
                Cursor.Current = currentCursor;
                if (exp.HResult == -2147024864)
                {
                    MessageBox.Show(this, "\"" + this.saveFileDialog1.FileName + "\" is already is use.\n\nPlease close it and generate report again.");
                }
                else
                if (exp.HResult == -2147024891)
                {
                    MessageBox.Show(this, "You did not have rights to save file on selected location.\n\nPlease run as administrator.");
                }
                else
                {
                    MessageBox.Show(this, exp.Message);
                }

            }
        }

        private void SaveAsExcel(Dictionary<string, List<CardHolderReportInfo>> data, string sheetName, string heading)
        {
            Cursor currentCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (data != null)
                {



                    using (ExcelPackage ex = new ExcelPackage())
                    {
                        ExcelWorksheet work = ex.Workbook.Worksheets.Add(sheetName);

                        work.View.ShowGridLines = false;
                        work.Cells.Style.Font.Name = "Segoe UI Light";

                        work.Column(1).Width = 20;
                        work.Column(2).Width = 54;
                        work.Column(3).Width = 25;
                        work.Column(4).Width = 25;
                        work.Column(5).Width = 25;
                        work.Column(6).Width = 25;

                        //Heading
                        work.Cells["A1:B2"].Merge = true;
                        work.Cells["A1:B2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        work.Cells["A1:B2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 213, 180));
                        work.Cells["A1:B2"].Style.Font.Size = 22;
                        work.Cells["A1:B2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells["A1:B2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        //work.Cells["A1:B2"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells["A1:B2"].Value = heading;

                        // img variable actually is your image path
                        System.Drawing.Image myImage = System.Drawing.Image.FromFile("Images/logo.png");

                        var pic = work.Drawings.AddPicture("Logo", myImage);

                        pic.SetPosition(5, 600);

                        int row = 4;


                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Report Time: ";
                        work.Cells[row, 2].Value = DateTime.Now.ToString();
                        work.Row(row).Height = 20;

                        row++;
                        row++;
                        //Sections and Data



                        foreach (KeyValuePair<string, List<CardHolderReportInfo>> category in data)
                        {

                            if (category.Value == null)
                            {
                                continue;
                            }

                            //Section
                            work.Cells[row, 1].Style.Font.Bold = true;
                            work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            work.Cells[row, 1].Value = "Category:";
                            work.Cells[row, 2].Value = category.Key;
                            work.Row(row).Height = 20;

                            //Data
                            row++;


                            work.Cells[row, 1, row, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            work.Cells[row, 1, row, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            work.Cells[row, 1, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            work.Cells[row, 1, row, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            
                            work.Cells[row, 1, row, 6].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 6].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 6].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 6].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                          
                            work.Cells[row, 1, row, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            work.Cells[row, 1, row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(253, 233, 217));
                            work.Cells[row, 1, row, 6].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            work.Cells[row, 1, row, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                           

                            work.Cells[row, 1].Value = "Category";
                            work.Cells[row, 2].Value = "First Name";
                            work.Cells[row, 3].Value = "CNIC Number";
                            work.Cells[row, 4].Value = "Blocked Status";
                            work.Cells[row, 5].Value = "Company";
                            work.Cells[row, 6].Value = "Total Days";
                            work.Row(row).Height = 20;

                            for (int i = 0; i < category.Value.Count; i++)
                            {
                                row++;
                                work.Cells[row, 1, row, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1, row, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1, row, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                work.Cells[row, 1, row, 6].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 6].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 6].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 6].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));             

                                if (i % 2 == 0)
                                {
                                    work.Cells[row, 1, row, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    work.Cells[row, 1, row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                }

                                work.Cells[row, 1, row, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                CardHolderReportInfo chl = category.Value[i];


                               
                                work.Cells[row, 1].Value = !Object.ReferenceEquals(null, chl.Category) ? chl.Category.ToString() : "Not Defined";
                                work.Cells[row, 2].Value = !Object.ReferenceEquals(null, chl.FirstName) ? chl.FirstName.ToString() : "Not Defined";
                                work.Cells[row, 3].Value = !Object.ReferenceEquals(null, chl.CNICNumber) ? chl.CNICNumber.ToString() : "Not Defined";
                                work.Cells[row, 4].Value = !Object.ReferenceEquals(null, chl.BlockedStatus) ? chl.BlockedStatus.ToString() : "Not Defined";
                                work.Cells[row, 5].Value = !Object.ReferenceEquals(null, chl.Company) ? chl.Company.ToString() : "Not Defined";
                                work.Cells[row, 6].Value = !Object.ReferenceEquals(null, chl.TotalDaysCkekin) ? chl.TotalDaysCkekin.ToString() : "Not Defined";





                                work.Row(row).Height = 20;
                            }

                            row++;
                            row++;
                        }


                        ex.SaveAs(new System.IO.FileInfo(this.saveFileDialog1.FileName));

                        System.Diagnostics.Process.Start(this.saveFileDialog1.FileName);
                    }
                }
                Cursor.Current = currentCursor;
            }
            catch (Exception exp)
            {
                Cursor.Current = currentCursor;
                if (exp.InnerException != null && exp.InnerException.InnerException != null)
                {
                    if (exp.InnerException.InnerException.HResult == -2147024864)
                    {
                        MessageBox.Show(this, "\"" + this.saveFileDialog1.FileName + "\" is already is use.\n\nPlease close it and generate report again.");
                    }
                    if (exp.InnerException.InnerException.HResult == -2147024891)
                    {
                        MessageBox.Show(this, "You did not have rights to save file on selected location.\n\nPlease run as administrator.");
                    }
                }
                else
                {
                    MessageBox.Show(this, exp.Message);
                }

            }

        }

        private void AddMainHeading(Table table, string heading)
        {
            Cell headingCell = new Cell(2, 4);
            headingCell.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            headingCell.SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3));
            headingCell.Add(new Paragraph(heading).SetFontSize(22F).SetBackgroundColor(new DeviceRgb(252, 213, 180))
                // .SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 3))
                );
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(iText.IO.Image.ImageDataFactory.Create("Images/logo.png"));

            table.AddCell(headingCell);
          //  table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(img).SetMarginLeft(60F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
        }

        private void AddNewEmptyRow(Table table, bool removeBottomBorder = true)
        {
            table.StartNewRow();

            if (removeBottomBorder)
            {
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            }
            else
            {
                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

              

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            }

        }

        private void AddCategoryRow(Table table, string categoryName)
        {
            table.StartNewRow();

            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Category:").
                    SetFontSize(11F).
                    SetBold().
                    SetFontColor(new DeviceRgb(247, 150, 70))).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(categoryName).
                    SetFontSize(11F).SetUnderline()).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
           
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
        }
        
        
        private void AddTableHeaderRow(Table table)
        {
            table.StartNewRow();

           
            table.AddCell(new Cell().
                    Add(new Paragraph("Category").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("First Name").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("CNIC Number").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Blocked Status").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Company").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Total Days").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph("CNIC Number").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph("Company Name").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }

        private void AddTableDataRow(Table table, CardHolderReportInfo chl, bool altRow)
        {
            if (chl == null)
            {
                return;
            }



            table.StartNewRow();

           
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.Category) ? string.Empty : chl.Category).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.FirstName) ? string.Empty : chl.FirstName).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.CNICNumber) ? string.Empty : chl.CNICNumber).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.BlockedStatus) ? string.Empty : chl.BlockedStatus).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                  Add(new Paragraph(string.IsNullOrEmpty(chl.Company) ? string.Empty : chl.Company).
                  SetFontSize(11F)).
              SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
              SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
              SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
              SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                 Add(new Paragraph(chl.TotalDaysCkekin.ToString()).
                 SetFontSize(11F)).
             SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
             SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
             SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
             SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.Crew) ? string.Empty : chl.Crew).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.Cadre) ? string.Empty : chl.Cadre).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.CNICNumber) ? string.Empty : chl.CNICNumber).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.Company) ? string.Empty : chl.Company).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }
    }
}
