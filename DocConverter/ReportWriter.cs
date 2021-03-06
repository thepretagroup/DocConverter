﻿using DocConverter.Properties;
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
        private const float TopMargin = 0.80f * 72;
        private const float BottomMargin = 0.75f * 72;

        private const bool INCLUDE_LOGO = false;

        private IList<Report> Reports { get; set; }

        public ReportWriter(IList<Report> reports)
        {
            Reports = reports;
        }

        #region DocX Writer
        public void CreateDocX(string docxOutputFile)
        {
            // Create a new document
            using (var document = DocX.Create(docxOutputFile))
            {
                document.MarginLeft = LeftRightMargin;
                document.MarginRight = LeftRightMargin;
                document.MarginTop = TopMargin;
                document.PageWidth = LetterPageWidth;
                document.PageHeight = LetterPageHeight;

                if (INCLUDE_LOGO)
                {
                    Include_Logo(document);
                }

                // Add a title
                document.InsertParagraph(Resources.ReportHeader).Font("Arial").FontSize(18)
                    .Bold().UnderlineStyle(UnderlineStyle.thick).Alignment = Alignment.left;

                WriteHeaderInfo(document, Reports[0]);
                document.InsertParagraph(string.Empty).Font("Arial").SpacingAfter(20);

                var table = CreateAnalysisTable(document);

                var firstReport = true;
                foreach (var report in Reports)
                {
                    if (!firstReport)
                    {
                        var row = table.InsertRow();
                        row.MergeCells(0, 3);
                        row.Cells[0].Paragraphs[0].Append("ExVivo Tar,Get Analysis??").Font("Arial").FontSize(14)
                            .Bold().UnderlineColor(System.Drawing.Color.Black);
                    }
                    if (report.DrugEffects.Count > 0)
                    {
                        WriteDrugEffects(table, Resources.SingleDoseEffectHeader, report.DrugEffects);
                    }
                    if (report.MultiDrugEffects.Count > 0)
                    {
                        WriteDrugEffects(table, Resources.MultiDoseEffectHeader, report.MultiDrugEffects.ToList<DrugEffect>());
                    }
                    firstReport = false;
                }
                document.InsertTable(table);

                WriteInterpretation(document, Reports[0]);

                WriteExVivoBest(document);

                WriteSignature(document, Reports[0]);

                document.Save();
            }
        }

        private void Include_Logo(DocX document)
        {
            using (var imageStream = new MemoryStream())
            {
                Resources.Logo_Nagourney_Cancer_Institute.Save(imageStream, ImageFormat.Jpeg);
                imageStream.Position = 0;
                var image = document.AddImage(imageStream);
                var picture = image.CreatePicture(
                    Resources.Logo_Nagourney_Cancer_Institute.Height / 8,
                    Resources.Logo_Nagourney_Cancer_Institute.Width / 8);
                //picture.WrappingStyle = PictureWrapText.WrapBehindText;  // Not available w/ Free version

                document.AddHeaders();
                document.Headers.Odd.Paragraphs[0].Font("Arial").AppendPicture(picture).Alignment = Alignment.right;

            }
        }

private void WriteExVivoBest(DocX document)
        {
            document.InsertParagraph(string.Empty).Font("Arial").SpacingAfter(24);

            document.InsertParagraph(Resources.ExVivoHeader).Font("Arial").FontSize(12).Bold().Alignment = Alignment.left;
        }

        private void WriteInterpretation(DocX document, Report report)
        {
            document.InsertParagraph(string.Empty).Font("Arial").SpacingAfter(24);

            document.InsertParagraph("DATA ANALYSIS:").Font("Arial").FontSize(14).Bold().Alignment = Alignment.left;

            document.InsertParagraph(Resources.DataAnalysis).Font("Arial").FontSize(10);
            document.InsertParagraph(string.Empty).Font("Arial");
            document.InsertParagraph(Resources.DataAnalysisNote).Font("Arial").FontSize(10);
        }

        private void WriteSignature(DocX document, Report report)
        {
            document.InsertParagraph(string.Empty).Font("Arial").SpacingAfter(24);

            document.InsertParagraph(Resources.SignatureBlock).Font("Arial").FontSize(12);
        }

        private void WriteDrugEffects(Table table, string header, List<DrugEffect> drugEffects)
        {
            var row = table.InsertRow();
            row.MergeCells(0, 3);
            row.Cells[0].Paragraphs[0].Append(header).Font("Arial").FontSize(14).Bold().Alignment = Alignment.center;

            foreach (var drugEffect in drugEffects
                .OrderBy(de => de.ExVivo.Rank)
                .ThenBy(de => (de is MultiDrugEffect) ? (de as MultiDrugEffect).ExVivoSynergy.Rank : 1)
                .ThenBy(de => de.Drug))
            {
                row = table.InsertRow();
                row.Cells[0].Paragraphs[0].Append(drugEffect.Drug).Font("Arial").FontSize(10d);
                var exVivoActivity = row.Cells[1].Paragraphs[0].Append(drugEffect.ExVivo.Activity).Font("Arial").FontSize(10d);
                if (drugEffect.ExVivo.Interpretation.Equals("Higher"))
                {
                    BoldItalicize(exVivoActivity);
                }
                if (drugEffect is MultiDrugEffect)
                {
                    row.Cells[2].Paragraphs[0].Append((drugEffect as MultiDrugEffect).ExVivoSynergy.Synergy).Font("Arial").FontSize(10d);
                }
                var exVivoInterpretation = row.Cells[3].Paragraphs[0].Append(drugEffect.ExVivo.Interpretation).Font("Arial").FontSize(10d);
                if (drugEffect.ExVivo.Interpretation.Equals("Higher"))
                {
                    BoldItalicize(exVivoInterpretation);
                }
            }
        }

        private void BoldItalicize(Paragraph paragraph)
        {
            paragraph.Bold().Italic();
        }

        private Table CreateAnalysisTable(DocX document)
        {
            var table = document.AddTable(4, 4);

            table.Design = TableDesign.None;
            table.AutoFit = AutoFit.ColumnWidth;
            table.SetWidths(new[] { 225f, 94f, 80f, 135f });

            table.Rows[0].Cells[1].Paragraphs[0].Append("Ex Vivo").Font("Arial").FontSize(10d).Bold();
            table.Rows[0].Cells[2].Paragraphs[0].Append("Ex Vivo").Font("Arial").FontSize(10d).Bold();
            table.Rows[0].Cells[3].Paragraphs[0].Append("Ex Vivo").Font("Arial").FontSize(10d).Bold();
            table.Rows[1].Cells[1].Paragraphs[0].Append("Activity").Font("Arial").FontSize(10d).Bold();
            table.Rows[1].Cells[2].Paragraphs[0].Append("Synergy").Font("Arial").FontSize(10d).Bold();
            table.Rows[1].Cells[3].Paragraphs[0].Append("Interpretation").Font("Arial").FontSize(10d).Bold();
            table.Rows[2].Cells[0].Paragraphs[0].Append("Drug").Font("Arial").FontSize(10).Bold();
            table.Rows[2].Cells[3].Paragraphs[0].Append("Response Expectation").Font("Arial").FontSize(10d).Bold();
            table.Rows[3].Cells[3].Paragraphs[0].Append("Compared with Database").Font("Arial").FontSize(10d).Bold();

            var border = new Border(BorderStyle.Tcbs_thick, BorderSize.six, 0, System.Drawing.Color.Black);
            foreach (var cell in table.Rows[3].Cells)
            {
                cell.SetBorder(TableCellBorderType.Bottom, border);
            }

            return table;
        }

        private void WriteHeaderInfo(DocX document, Report report)
        {
            document.InsertParagraph(string.Empty).Font("Arial").SpacingAfter(30);

            WriteSpecLine(document, "Patient:", report.Patient, "Assay Date:", report.AssayDate);
            WriteSpecLine(document, "DOB:", string.Empty, "Report Date:", report.ReportDate);
            WriteSpecLine(document, "Dx:", report.Dx, "Specimen Number:", report.SpecimenNumber);

            var physicianTitle = "Physician:";
            foreach (var physician in report.Physician)
            {
                WriteSpecLine(document, physicianTitle, physician);
                physicianTitle = string.Empty;
            }
        }

        private void WriteSpecLine(DocX document, string leftName, string leftValue, string rightName = "", string rightValue = "")
        {
            var paragraph = document.InsertParagraph(leftName)
                .Font("Arial")
                .FontSize(10d)
                .Bold()
                .InsertTabStopPosition(Alignment.left, 100)
                .InsertTabStopPosition(Alignment.left, 300)
                .InsertTabStopPosition(Alignment.left, 400);

            paragraph.Append("\t" + leftValue).Font("Arial").FontSize(10d);
            paragraph.Append("\t" + rightName).Font("Arial").FontSize(10d).Bold();
            paragraph.Append("\t" + rightValue).Font("Arial").FontSize(10d);
            paragraph.SetLineSpacing(LineSpacingType.Line, 1.5f);
        }
        #endregion
    }
}