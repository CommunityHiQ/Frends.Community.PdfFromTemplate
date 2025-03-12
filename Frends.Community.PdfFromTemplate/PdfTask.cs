using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using Newtonsoft.Json;
using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout; // CORRECTED USING STATEMENT
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using SimpleImpersonation;
using SkiaSharp;
using Path = System.IO.Path;
using iText.IO.Font.Constants;
using iText.Kernel.Events;
using iText.Kernel.Pdf.Canvas;


namespace Frends.Community.PdfFromTemplate
{
    public class PdfTask
    {
        public static Output CreatePdf([PropertyTab] FileProperties outputFile, [PropertyTab] DocumentContent content, [PropertyTab] Options options)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                var docContent = JsonConvert.DeserializeObject<DocumentDefinition>(content.ContentJson);
                if (docContent == null)
                {
                    throw new ArgumentException("Content JSON could not be deserialized to DocumentDefinition.", nameof(content.ContentJson));
                }

                using var memoryStream = new MemoryStream();
                using var writer = new PdfWriter(memoryStream);
                using var pdf = new PdfDocument(writer);
                using var document = new Document(pdf);

                // Set document margins
                document.SetMargins(
                    (float)UnitConverter.ConvertCmToPoint((float)docContent.MarginTopInCm),
                    (float)UnitConverter.ConvertCmToPoint((float)docContent.MarginRightInCm),
                    (float)UnitConverter.ConvertCmToPoint((float)docContent.MarginBottomInCm),
                    (float)UnitConverter.ConvertCmToPoint((float)docContent.MarginLeftInCm)
                );

                // Set document metadata
                pdf.GetDocumentInfo().SetTitle(docContent.Title);
                pdf.GetDocumentInfo().SetAuthor(docContent.Author);

                SetupPage(pdf, docContent);

                foreach (var pageElement in docContent.DocumentElements)
                {
                    ProcessDocumentElement(document, pageElement, pdf, docContent);
                }

                document.Close();

                byte[] pdfBytes = memoryStream.ToArray();
                string fullFilePath = Path.Combine(outputFile.Directory, outputFile.FileName);
                fullFilePath = HandleFileExists(outputFile, fullFilePath);

                if (outputFile.SaveToDisk)
                {
                    WritePdfToDisk(fullFilePath, pdfBytes, options);
                }

                return new Output { Success = true, FileName = fullFilePath, ResultAsByteArray = options.GetResultAsByteArray ? pdfBytes : null };
            }
            catch (Exception ex)
            {
                if (options.ThrowErrorOnFailure)
                {
                    throw;
                }
                return new Output { Success = false, ErrorMessage = ex.Message };
            }
        }

        private static void ProcessDocumentElement(Document document, DocumentElement element, PdfDocument pdf, DocumentDefinition docDefinition)
        {
            switch (element.ElementType)
            {
                case ElementTypeEnum.Paragraph:
                    AddParagraph(document, (ParagraphDefinition)element);
                    break;
                case ElementTypeEnum.Image:
                    AddImage(document, (ImageDefinition)element);
                    break;
                case ElementTypeEnum.Table:
                    AddTable(document, (TableDefinition)element, pdf);
                    break;
                case ElementTypeEnum.PageBreak:
                    document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                    SetupPage(pdf, docDefinition);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(element.ElementType), $"Unsupported element type: {element.ElementType}");
            }
        }

        private static void SetupPage(PdfDocument pdf, DocumentDefinition docDefinition)
        {
            PageSize pageSize = docDefinition.PageSize switch
            {
                PageSizeEnum.A0 => PageSize.A0,
                PageSizeEnum.A1 => PageSize.A1,
                PageSizeEnum.A2 => PageSize.A2,
                PageSizeEnum.A3 => PageSize.A3,
                PageSizeEnum.A4 => PageSize.A4,
                PageSizeEnum.A5 => PageSize.A5,
                PageSizeEnum.A6 => PageSize.A6,
                PageSizeEnum.B5 => PageSize.B5,
                PageSizeEnum.Letter => PageSize.LETTER,
                PageSizeEnum.Legal => PageSize.LEGAL,
                PageSizeEnum.Ledger => PageSize.LEDGER,
                _ => PageSize.A4,
            };

            if (docDefinition.PageOrientation == PageOrientationEnum.Landscape)
            {
                pageSize = pageSize.Rotate();
            }

            pdf.AddNewPage(pageSize);
        }

        private static void AddParagraph(Document document, ParagraphDefinition paragraphDef)
        {
            var paragraph = new Paragraph(paragraphDef.Text);
            ApplyStyleSettings(paragraph, paragraphDef.StyleSettings);

            if (paragraphDef.StyleSettings.BorderWidthInPt > 0)
            {
                var border = new SolidBorder((float)paragraphDef.StyleSettings.BorderWidthInPt);
                switch (paragraphDef.StyleSettings.BorderStyle)
                {
                    case BorderStyleEnum.All:
                        paragraph.SetBorder(border);
                        break;
                    case BorderStyleEnum.Top:
                        paragraph.SetBorderTop(border);
                        break;
                    case BorderStyleEnum.Bottom:
                        paragraph.SetBorderBottom(border);
                        break;
                    case BorderStyleEnum.None:
                        paragraph.SetBorder(Border.NO_BORDER);
                        break;
                }
            }

            document.Add(paragraph);
        }

        private static void AddImage(Document document, ImageDefinition imageDef)
        {
            var imagePath = imageDef.ImagePath.Replace("\\\\", "\\");
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Image file not found: {imagePath}", imagePath);
            }

            using var skBitmap = SKBitmap.Decode(imagePath);
            using var skImage = SKImage.FromBitmap(skBitmap);
            using var skData = skImage.Encode(SKEncodedImageFormat.Png, 100);
            byte[] imageBytes = skData.ToArray();

            var imageData = ImageDataFactory.Create(imageBytes);
            var itextImage = new Image(imageData);

            if (imageDef.ImageWidthInCm > 0)
            {
                itextImage.ScaleToFit((float)UnitConverter.ConvertCmToPoint((float)imageDef.ImageWidthInCm), float.MaxValue);

                if (imageDef.ImageHeightInCm > 0 && !imageDef.LockAspectRatio)
                {
                    itextImage.ScaleAbsolute((float)UnitConverter.ConvertCmToPoint((float)imageDef.ImageWidthInCm),
                                           (float)UnitConverter.ConvertCmToPoint((float)imageDef.ImageHeightInCm));
                }
            }

            SetHorizontalAlignment(itextImage, imageDef.Alignment.ToString());
            document.Add(itextImage);
        }

        private static void AddTable(Document document, TableDefinition tableDef, PdfDocument pdf)
        {
            float[] columnWidths = tableDef.Columns.Select(c => (float)c.WidthInCm).ToArray();
            float totalWidth = columnWidths.Sum();
            float pageWidth = pdf.GetDefaultPageSize().GetWidth() / UnitConverter.ConvertCmToPoint(1);
            float availableWidth = pageWidth - document.GetLeftMargin() / UnitConverter.ConvertCmToPoint(1) - document.GetRightMargin() / UnitConverter.ConvertCmToPoint(1);

            if (totalWidth > availableWidth)
            {
                throw new Exception($"Table width ({totalWidth:F1} cm) exceeds available page width ({availableWidth:F1} cm)");
            }

            // Create main table
            var table = new Table(UnitValue.CreatePointArray(columnWidths.Select(w => UnitConverter.ConvertCmToPoint(w)).ToArray()));
            table.SetWidth(UnitValue.CreatePercentValue(100));

            // Add header row
            if (tableDef.HasHeaderRow)
            {
                if (tableDef.HeaderData != null && tableDef.HeaderData.Any())
                {
                    // Use custom header data
                    AddTableRow(table, tableDef.HeaderData[0].Select(kvp => new TableCellData { Text = kvp.Value }).ToList(), tableDef.Columns, tableDef.StyleSettings, tableDef.TableType == TableTypeEnum.Header);
                }
                else
                {
                    // Use column names
                    foreach (var column in tableDef.Columns)
                    {
                        var cell = new Cell().Add(new Paragraph(column.Name));
                        if (tableDef.TableType == TableTypeEnum.Header)
                        {
                            cell.SetBorder(Border.NO_BORDER);
                        }
                        else
                        {
                            // Add border for regular tables
                            cell.SetBorder(new SolidBorder(0.5f));
                        }
                        ApplyStyleSettings(cell, tableDef.StyleSettings);
                        table.AddHeaderCell(cell);
                    }
                }
            }

            // Add data rows
            foreach (var dataRow in tableDef.RowData)
            {
                // Skip if this is a duplicate of the header row data (if using column names as header)
                if (tableDef.HasHeaderRow && tableDef.HeaderData == null && dataRow.Values.SequenceEqual(tableDef.Columns.Select(c => c.Name)))
                {
                    continue;
                }
                AddTableRow(table, dataRow.Select(kvp => new TableCellData { Text = kvp.Value }).ToList(), tableDef.Columns, tableDef.StyleSettings, tableDef.TableType == TableTypeEnum.Header);
            }

            // Apply table borders based on table type
            if (tableDef.TableType == TableTypeEnum.Header)
            {
                var headerHandler = new TableHeaderEventHandler(table);
                pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, headerHandler);
            }
            else if (tableDef.TableType == TableTypeEnum.Footer)
            {
                var footerHandler = new TableFooterEventHandler(table);
                pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, footerHandler);
            }
            else
            {
                document.Add(table);
            }
        }

        private static void AddTableRow(Table table, List<TableCellData> rowData, List<ColumnDefinition> columns, StyleSettingsDefinition styleSettings, bool isHeaderTable = false)
        {
            for (int i = 0; i < columns.Count; i++)
            {
                var columnDef = columns[i];
                var cellData = rowData[i];

                Cell cell;
                switch (columnDef.Type)
                {
                    case ColumnTypeEnum.Text:
                        cell = new Cell().Add(new Paragraph(cellData.Text ?? ""));
                        if (isHeaderTable)
                        {
                            cell.SetBorder(Border.NO_BORDER);
                        }
                        else
                        {
                            cell.SetBorder(new SolidBorder(0.5f));
                        }
                        ApplyStyleSettings(cell, styleSettings);
                        break;
                    case ColumnTypeEnum.Image:
                        cell = CreateImageCell(cellData.ImagePath ?? cellData.Text, (float)columns[i].WidthInCm);
                        if (isHeaderTable)
                        {
                            cell.SetBorder(Border.NO_BORDER);
                        }
                        else
                        {
                            cell.SetBorder(new SolidBorder(0.5f));
                        }
                        break;
                    case ColumnTypeEnum.PageNum:
                        cell = new Cell().Add(new Paragraph(new Text("")));
                        if (isHeaderTable)
                        {
                            cell.SetBorder(Border.NO_BORDER);
                        }
                        else
                        {
                            cell.SetBorder(new SolidBorder(0.5f));
                        }
                        ApplyStyleSettings(cell, styleSettings);
                        break;
                    default:
                        cell = new Cell().Add(new Paragraph(""));
                        if (isHeaderTable)
                        {
                            cell.SetBorder(Border.NO_BORDER);
                        }
                        else
                        {
                            cell.SetBorder(new SolidBorder(0.5f));
                        }
                        break;
                }
                table.AddCell(cell);
            }
        }

        private static Cell CreateImageCell(string imagePath, float columnWidthInCm)
        {
            if (string.IsNullOrEmpty(imagePath))
            {
                throw new FileNotFoundException("Image path cannot be empty");
            }

            var cell = new Cell();
            imagePath = imagePath.Replace("\\\\", "\\");
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Image file not found: {imagePath}", imagePath);
            }

            using var skBitmap = SKBitmap.Decode(imagePath);
            using var skImage = SKImage.FromBitmap(skBitmap);
            using var skData = skImage.Encode(SKEncodedImageFormat.Png, 100);
            byte[] imageBytes = skData.ToArray();

            var imageData = ImageDataFactory.Create(imageBytes);
            var image = new Image(imageData);
            image.ScaleToFit(UnitConverter.ConvertCmToPoint(columnWidthInCm), float.MaxValue);
            cell.Add(image);
            return cell;
        }

        private static void ApplyStyleSettings(IBlockElement element, StyleSettingsDefinition settings)
        {
            // Font
            if (!string.IsNullOrEmpty(settings.FontFamily))
            {
                SetFont(element, settings.FontFamily);
            }

            if (element is Paragraph paragraph)
            {
                paragraph.SetFontSize((float)settings.FontSizeInPt)
                        .SetTextAlignment(GetTextAlignment(settings.HorizontalAlignment))
                        .SetFixedLeading((float)settings.LineSpacingInPt)
                        .SetMarginTop((float)settings.SpacingBeforeInPt)
                        .SetMarginBottom((float)settings.SpacingAfterInPt);

                ApplyFontStyle(paragraph, settings.FontStyle);
            }
            else if (element is Cell cell)
            {
                cell.SetFontSize((float)settings.FontSizeInPt)
                    .SetTextAlignment(GetTextAlignment(settings.HorizontalAlignment))
                    .SetVerticalAlignment(GetVerticalAlignment(settings.VerticalAlignment));

                ApplyFontStyle(cell, settings.FontStyle);
            }
        }

        private static void SetFont(IBlockElement element, string fontFamily)
        {
            try
            {
                PdfFont pdfFont;
                if (fontFamily.Equals("Times New Roman", StringComparison.OrdinalIgnoreCase))
                {
                    pdfFont = PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN);
                }
                else if (fontFamily.Equals("Helvetica", StringComparison.OrdinalIgnoreCase))
                {
                    pdfFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);
                }
                else
                {
                    var fontProgram = FontProgramFactory.CreateFont(fontFamily);
                    pdfFont = PdfFontFactory.CreateFont(fontProgram, PdfEncodings.WINANSI, PdfFontFactory.EmbeddingStrategy.FORCE_EMBEDDED);
                }

                if (element is Paragraph paragraph)
                {
                    paragraph.SetFont(pdfFont);
                }
                else if (element is Cell cell)
                {
                    cell.SetFont(pdfFont);
                }
            }
            catch
            {
                var pdfFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

                if (element is Paragraph paragraph)
                {
                    paragraph.SetFont(pdfFont);
                }
                else if (element is Cell cell)
                {
                    cell.SetFont(pdfFont);
                }
            }
        }

        private static void ApplyFontStyle(IBlockElement element, FontStyleEnum fontStyle)
        {
            if (element is Paragraph paragraph)
            {
                switch (fontStyle)
                {
                    case FontStyleEnum.Bold:
                        paragraph.SetBold();
                        break;
                    case FontStyleEnum.Italic:
                        paragraph.SetItalic();
                        break;
                    case FontStyleEnum.BoldItalic:
                        paragraph.SetBold().SetItalic();
                        break;
                    case FontStyleEnum.Underline:
                        paragraph.SetUnderline();
                        break;
                }
            }
            else if (element is Cell cell)
            {
                switch (fontStyle)
                {
                    case FontStyleEnum.Bold:
                        cell.SetBold();
                        break;
                    case FontStyleEnum.Italic:
                        cell.SetItalic();
                        break;
                    case FontStyleEnum.BoldItalic:
                        cell.SetBold().SetItalic();
                        break;
                    case FontStyleEnum.Underline:
                        cell.SetUnderline();
                        break;
                }
            }
        }

        private static TextAlignment GetTextAlignment(HorizontalAlignmentEnum alignment)
        {
            return alignment switch
            {
                HorizontalAlignmentEnum.Left => TextAlignment.LEFT,
                HorizontalAlignmentEnum.Center => TextAlignment.CENTER,
                HorizontalAlignmentEnum.Right => TextAlignment.RIGHT,
                HorizontalAlignmentEnum.Justify => TextAlignment.JUSTIFIED,
                _ => TextAlignment.LEFT
            };
        }

        private static VerticalAlignment GetVerticalAlignment(VerticalAlignmentEnum alignment)
        {
            return alignment switch
            {
                VerticalAlignmentEnum.Top => VerticalAlignment.TOP,
                VerticalAlignmentEnum.Center => VerticalAlignment.MIDDLE,
                VerticalAlignmentEnum.Bottom => VerticalAlignment.BOTTOM,
                _ => VerticalAlignment.BOTTOM
            };
        }

        private static void ApplyTableBorders(Table table, StyleSettingsDefinition settings)
        {
            if (settings.BorderWidthInPt > 0)
            {
                var border = new SolidBorder((float)settings.BorderWidthInPt);
                switch (settings.BorderStyle)
                {
                    case BorderStyleEnum.All:
                        table.SetBorder(border);
                        break;
                    case BorderStyleEnum.Top:
                        table.SetBorderTop(border);
                        break;
                    case BorderStyleEnum.Bottom:
                        table.SetBorderBottom(border);
                        break;
                    case BorderStyleEnum.None:
                        table.SetBorder(Border.NO_BORDER);
                        break;
                }
            }
            else
            {
                table.SetBorder(Border.NO_BORDER);
            }
        }

        private static void SetHorizontalAlignment(IPropertyContainer element, string alignment)
        {
            element.SetProperty(Property.HORIZONTAL_ALIGNMENT, (HorizontalAlignment)ParseEnum<HorizontalAlignmentEnum>(alignment));
        }

        private static string HandleFileExists(FileProperties outputFile, string fullFilePath)
        {
            if (File.Exists(fullFilePath))
            {
                if (outputFile.FileExistsAction == FileExistsActionEnum.Error)
                {
                    throw new Exception($"File {fullFilePath} already exists.");
                }
                else if (outputFile.FileExistsAction == FileExistsActionEnum.Rename)
                {
                    var directory = Path.GetDirectoryName(fullFilePath);
                    var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fullFilePath);
                    var extension = Path.GetExtension(fullFilePath);
                    var counter = 1;

                    while (File.Exists(fullFilePath))
                    {
                        fullFilePath = Path.Combine(directory, $"{fileNameWithoutExtension}_({counter}){extension}");
                        counter++;
                    }
                }
            }
            return fullFilePath;
        }

        private static void WritePdfToDisk(string fullFilePath, byte[] pdfBytes, Options options)
        {
            if (!options.UseGivenCredentials)
            {
                File.WriteAllBytes(fullFilePath, pdfBytes);
                return;
            }

            if (!OperatingSystem.IsWindows())
            {
                Console.WriteLine("Warning: Impersonation is only supported on Windows. Writing file as current user.");
                File.WriteAllBytes(fullFilePath, pdfBytes);
                return;
            }

            var domainAndUserName = GetDomainAndUserName(options.UserName);
            var credentials = new UserCredentials(domainAndUserName[0], domainAndUserName[1], options.Password);
            using (var userContext = credentials.LogonUser(LogonType.NewCredentials))
            {
                WindowsIdentity.RunImpersonated(userContext, () =>
                {
                    File.WriteAllBytes(fullFilePath, pdfBytes);
                });
            }
        }

        private static string[] GetDomainAndUserName(string userName)
        {
            var parts = userName.Split('\\');
            if (parts.Length != 2)
            {
                throw new ArgumentException($"UserName field must be of format domain\\username was: {userName}");
            }
            return parts;
        }

        private static T ParseEnum<T>(string value) where T : struct, Enum
        {
            if (Enum.TryParse(value, true, out T result))
            {
                return result;
            }
            //Consider logging the enum name here
            throw new ArgumentException($"Invalid value '{value}' for enum type '{typeof(T).Name}'.  Valid values are: {string.Join(", ", Enum.GetNames(typeof(T)))}");
        }

        public class TableCellData
        {
            public string Text { get; set; }
            public string ImagePath { get; set; }
            public StyleSettingsDefinition Style { get; set; }
        }
    }

    // --- TableHeaderEventHandler ---
    public class TableHeaderEventHandler : IEventHandler
    {
        private readonly Table _table;

        public TableHeaderEventHandler(Table table)
        {
            _table = table;
            // Remove all borders from the table
            _table.SetBorder(Border.NO_BORDER);
            // Remove all cell borders
            foreach (var cell in _table.GetChildren())
            {
                if (cell is Cell tableCell)
                {
                    tableCell.SetBorder(Border.NO_BORDER);
                }
            }
        }

        public void HandleEvent(Event @event)
        {
            PdfDocumentEvent docEvent = (PdfDocumentEvent)@event;
            PdfDocument pdfDoc = docEvent.GetDocument();
            PdfPage page = docEvent.GetPage();
            Rectangle pageSize = page.GetPageSize();
            PdfCanvas pdfCanvas = new PdfCanvas(page.NewContentStreamBefore(), page.GetResources(), pdfDoc);

            // Get document margins in points (assuming standard A4 margins of 2.5cm from ModelDocument.json)
            float marginLeft = UnitConverter.ConvertCmToPoint(2.5f);
            float marginTop = UnitConverter.ConvertCmToPoint(2.5f);
            
            // Calculate positions based on document margins
            float x = marginLeft;
            float y = pageSize.GetTop() - marginTop;
            float width = pageSize.GetWidth() - (2 * marginLeft);

            // Use SetFixedPosition for reliable positioning
            _table.SetFixedPosition(x, y, width);

            // Use a Canvas to add the table to the page
            var canvas = new Canvas(pdfCanvas, pageSize);
            canvas.Add(_table);
            canvas.Close();
            pdfCanvas.Release();
        }
    }

    // --- TableFooterEventHandler ---
    public class TableFooterEventHandler : IEventHandler
    {
        private readonly Table _table;

        public TableFooterEventHandler(Table table)
        {
            _table = table;
            // Remove all borders from the table
            _table.SetBorder(Border.NO_BORDER);
            // Remove all cell borders
            foreach (var cell in _table.GetChildren())
            {
                if (cell is Cell tableCell)
                {
                    tableCell.SetBorder(Border.NO_BORDER);
                }
            }
        }

        public void HandleEvent(Event @event)
        {
            PdfDocumentEvent docEvent = (PdfDocumentEvent)@event;
            PdfDocument pdfDoc = docEvent.GetDocument();
            PdfPage page = docEvent.GetPage();
            Rectangle pageSize = page.GetPageSize();
            PdfCanvas pdfCanvas = new PdfCanvas(page.NewContentStreamBefore(), page.GetResources(), pdfDoc);

            // Define margins and positions
            float margin = 36;
            float x = margin;
            float y = margin; // Bottom of the page
            float width = pageSize.GetWidth() - (2 * margin); // Page width minus left and right margins

            // Draw page number (optional, but good practice for footers)
            int pageNumber = pdfDoc.GetPageNumber(page);
            pdfCanvas.BeginText()
                .SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 10)
                .MoveText(pageSize.GetWidth() / 2 - 20, y) // Center the page number
                .ShowText($"Page {pageNumber}")
                .EndText();

            // Use SetFixedPosition for reliable positioning
            _table.SetFixedPosition(x, y + 15, width); // Position the table *above* the page number

            // Use a Canvas to add the table
            var canvas = new Canvas(pdfCanvas, pageSize);
            canvas.Add(_table);
            canvas.Close();

            pdfCanvas.Release();
        }
    }


    public static class UnitConverter
    {
        public static float ConvertCmToPoint(float cm)
        {
            return cm * 28.3465f;
        }
    }
}