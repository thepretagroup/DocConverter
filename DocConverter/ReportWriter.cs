﻿using DocConverter.Properties;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using Xceed.Words.NET;

namespace DocConverter
{
    public class ReportWriter
    {
        private const float LetterPageWidth = 8.5f * 72;
        private const float LetterPageHeight = 11f * 72;
        private const float LeftRightMargin = 0.5f * 72;
        private const float TopBottomMargin = 0.75f * 72;

        private IList<Report> Reports { get; set; }

        public ReportWriter(IList<Report> reports)
        {
            Reports = reports;
        }

        #region DocX Writer
        public void CreateDocX(string docxOutputFile)
        {
            // Create a new document.
            using (var document = DocX.Create(docxOutputFile))
            {
                document.MarginLeft = LeftRightMargin;
                document.MarginRight = LeftRightMargin;
                document.MarginTop = TopBottomMargin;
                document.PageWidth = LetterPageWidth;
                document.PageHeight = LetterPageHeight;

                using (var imageStream = new MemoryStream())
                {
                    Resources.Logo_Nagourney_Cancer_Institute.Save(imageStream, ImageFormat.Jpeg);
                    imageStream.Position = 0;
                    var image = document.AddImage(imageStream);
                    var picture = image.CreatePicture(
                        Resources.Logo_Nagourney_Cancer_Institute.Height / 4,
                        Resources.Logo_Nagourney_Cancer_Institute.Width / 4);
                    //picture.WrappingStyle = PictureWrapText.WrapBehindText;  // Not available w/ Free version
                    document.InsertParagraph().AppendPicture(picture).Alignment = Alignment.right;
                }

                // Add a title
                document.InsertParagraph(Resources.ReportHeader).Font("Arial").FontSize(18).Bold().UnderlineStyle(UnderlineStyle.thick).Alignment = Alignment.left;

                WriteHeaderInfo(document, Reports[0]);
                document.InsertParagraph().SpacingAfter(20);

                var table = CreateAnalysisTable(document);

                var firstReport = true;
                foreach (var report in Reports)
                {
                    if (!firstReport)
                    {
                        var row = table.InsertRow();
                        row.MergeCells(0, 3);
                        row.Cells[0].Paragraphs[0].Append("ExVivo Tar,Get Analysis??").Font("Arial").FontSize(14).Bold().UnderlineColor(System.Drawing.Color.Black);
                    }
                    if (report.DrugEffects.Count > 0)
                    {
                        WriteDrugEffects(table, Resources.SingleDoseEffectHeader, report.DrugEffects);
                    }
                    if (report.MultiDrugEffects.Count > 0)
                    {
                        WriteDrugEffects(table, Resources.MultiDoseEffectHeader, report.MultiDrugEffects.ToList<DrugEffect>()); //  .Cast<IDrugEffect>().ToList());
                    }
                    firstReport = false;
                }
                document.InsertTable(table);

                WriteInterpretation(document, Reports[0]);

                WriteExVivoBest(document);

                WriteSignature(document, Reports[0]);

                // Save this document to disk.
                document.Save();
            }
        }

        private void WriteExVivoBest(DocX document)
        {
            document.InsertParagraph().SpacingAfter(24);

            document.InsertParagraph(Resources.ExVivoHeader).Font("Arial").FontSize(12).Bold().Alignment = Alignment.left;
        }

        private void WriteInterpretation(DocX document, Report report)
        {
            document.InsertParagraph().SpacingAfter(24);

            document.InsertParagraph("DATA ANALYSIS:").Font("Arial").FontSize(14).Bold().Alignment = Alignment.left;

            document.InsertParagraph(Resources.DataAnalysis).Font("Arial").FontSize(10);
            document.InsertParagraph();
            document.InsertParagraph(Resources.DataAnalysisNote ).Font("Arial").FontSize(10);
        }

        private void WriteSignature(DocX document, Report report)
        {
            document.InsertParagraph().SpacingAfter(24);
            foreach (var signatureLine in report.Signature)
            {
                document.InsertParagraph(signatureLine).Font("Arial").FontSize(12);
            }
        }

        private void WriteDrugEffects(Table table, string header, List<DrugEffect> drugEffects)
        {
            var row = table.InsertRow();
            row.MergeCells(0, 3);
            row.Cells[0].Paragraphs[0].Append(header).Font("Arial").FontSize(14).Bold().Alignment = Alignment.center;

            foreach (var drugEffect in drugEffects
                .FindAll(de => !de.Interpretation.Contains("Active"))
                .OrderBy(de => de.ExVivo.Rank)
                .ThenBy(de => de.Drug))
            {
                row = table.InsertRow();
                row.Cells[0].Paragraphs[0].Append(drugEffect.Drug).FontSize(10d);
                var exVivoActivity = row.Cells[1].Paragraphs[0].Append(drugEffect.ExVivo.Activity).FontSize(10d);
                if (drugEffect.ExVivo.Rank == 1)
                {
                    BoldItalicize(exVivoActivity);
                }
                if (drugEffect is MultiDrugEffect)
                {
                    row.Cells[2].Paragraphs[0].Append((drugEffect as MultiDrugEffect).Synergy).FontSize(10d);
                }
                var exVivoInterpretation = row.Cells[3].Paragraphs[0].Append(drugEffect.ExVivo.Interpretation).FontSize(10d);
                if (drugEffect.ExVivo.Rank == 1)
                {
                    BoldItalicize(exVivoInterpretation);
                }

            }
        }

        private void BoldItalicize(Paragraph paragraph)
        {
            paragraph.Bold().Italic();
        }

        private Table CreateAnalysisTable(DocX document) {
            var table = document.AddTable(4, 4);

            table.Design = TableDesign.None;
            table.AutoFit = AutoFit.ColumnWidth;
            table.SetWidths(new[] { 225f, 90f, 75f, 120f });

            table.Rows[0].Cells[1].Paragraphs[0].Append("Ex Vivo").FontSize(10d).Bold();
            table.Rows[0].Cells[2].Paragraphs[0].Append("Ex Vivo").FontSize(10d).Bold();
            table.Rows[0].Cells[3].Paragraphs[0].Append("Ex Vivo").FontSize(10d).Bold();
            table.Rows[1].Cells[1].Paragraphs[0].Append("Activity").FontSize(10d).Bold();
            table.Rows[1].Cells[2].Paragraphs[0].Append("Synergy").FontSize(10d).Bold();
            table.Rows[1].Cells[3].Paragraphs[0].Append("Interpretation").FontSize(10d).Bold();
            table.Rows[2].Cells[0].Paragraphs[0].Append("Drug").FontSize(10).Bold();
            table.Rows[2].Cells[3].Paragraphs[0].Append("Response Expectation").Font("Times Roman New").FontSize(10d);
            table.Rows[3].Cells[3].Paragraphs[0].Append("Compared with Database").Font("Times Roman New").FontSize(10d);

            var border = new Border(BorderStyle.Tcbs_thick, BorderSize.six, 0, System.Drawing.Color.Black);
            foreach (var cell in table.Rows[3].Cells)
            {
                cell.SetBorder(TableCellBorderType.Bottom, border);
            }

            return table;
        }

        private void WriteHeaderInfo(DocX document, Report report)
        {
            document.InsertParagraph().SpacingAfter(30);

            WriteSpecLine(document, "Patient:", report.Patient, "Assay Date:", report.AssayDate);
            WriteSpecLine(document, "Dx:", report.Dx, "Assay Quality:", "" /* report.AssayQuality */);
            WriteSpecLine(document, "Prior Rx:", report.PriorRx, "Report Date:", report.ReportDate);

            var physicianTitle = "Physician:";
            var specimenNumbertitle = "Specimen Number:";
            var specimenNumber = report.SpecimenNumber;
            foreach (var physician in report.Physician)
            {
                WriteSpecLine(document, physicianTitle, physician, specimenNumbertitle, specimenNumber);
                physicianTitle = specimenNumbertitle = specimenNumber = string.Empty;
            }
        }

        private void WriteSpecLine(DocX document, string leftName, string leftValue, string rightName, string rightValue)
        {
            var outputLine = string.Format("{0}\t{1}\t{2}\t{3}", leftName, leftValue, rightName, rightValue);
            Console.WriteLine("@@@@@@" + outputLine);

            var paragraph = document.InsertParagraph(outputLine)
                .FontSize(10d)
                .Bold()
                .InsertTabStopPosition(Alignment.left, 100)
                .InsertTabStopPosition(Alignment.left, 300)
                .InsertTabStopPosition(Alignment.left, 400);
        }
        #endregion
    }
}