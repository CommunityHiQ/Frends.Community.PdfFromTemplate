using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using Newtonsoft.Json;
using SimpleImpersonation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Frends.Community.PdfFromTemplate
{
    public class PdfTask
    {
        /// <summary>
        /// Creates PDF document from given content. See https://github.com/CommunityHiQ/Frends.Community.PdfFromTemplate
        /// </summary>
        /// <param name="outputFile"></param>
        /// <param name="content"></param>
        /// <param name="options"></param>
        /// <returns>Object { bool Success, string FileName, byte[] ResultAsByteArray }</returns>
        public static Output CreatePdf([PropertyTab]FileProperties outputFile,
            [PropertyTab]DocumentContent content,
            [PropertyTab]Options options)
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                DocumentDefinition docContent = JsonConvert.DeserializeObject<DocumentDefinition>(content.ContentJson);

                var document = new Document();
                if (!string.IsNullOrWhiteSpace(docContent.Title))
                {
                    document.Info.Title = docContent.Title;
                }
                if (!string.IsNullOrWhiteSpace(docContent.Author))
                {
                    document.Info.Author = docContent.Author;
                }

                Unit width, height;
                PageSetup.GetPageSize(docContent.PageSize.ConvertEnum<PageFormat>(), out width, out height);

                var section = document.AddSection();
                SetupPage(section.PageSetup, width, height, docContent);

                // index for stylename
                var elementNumber = 0;
                // add page elements
                foreach (var pageElement in docContent.DocumentElements)
                {
                    var styleName = $"style_{elementNumber}";
                    var style = document.Styles.AddStyle(styleName, "Normal");
                    switch (pageElement.ElementType)
                    {
                        case ElementTypeEnum.Paragraph:
                            SetFont(style, ((ParagraphDefinition)pageElement).StyleSettings);
                            SetParagraphStyle(style, ((ParagraphDefinition)pageElement).StyleSettings, false);
                            AddTextContent(section, ((ParagraphDefinition)pageElement).Text, style);
                            break;
                        case ElementTypeEnum.Image:
                            AddImage(section, (ImageDefinition)pageElement, width);
                            break;
                        case ElementTypeEnum.Table:
                            SetFont(style, ((TableDefinition)pageElement).StyleSettings);
                            SetParagraphStyle(style, ((TableDefinition)pageElement).StyleSettings, true);
                            AddTable(section, (TableDefinition)pageElement, width, style);
                            break;
                        case ElementTypeEnum.PageBreak:
                            section = document.AddSection();
                            SetupPage(section.PageSetup, width, height, docContent);
                            break;
                        default:
                            break;
                    }

                    ++elementNumber;

                }

                string fileName = Path.Combine(outputFile.Directory, outputFile.FileName);
                int fileNameIndex = 1;
                while (File.Exists(fileName) && outputFile.FileExistsAction != FileExistsActionEnum.Overwrite)
                {
                    switch (outputFile.FileExistsAction)
                    {
                        case FileExistsActionEnum.Error:
                            throw new Exception($"File {fileName} already exists.");
                        case FileExistsActionEnum.Rename:
                            fileName = Path.Combine(outputFile.Directory, $"{Path.GetFileNameWithoutExtension(outputFile.FileName)}_({fileNameIndex}){Path.GetExtension(outputFile.FileName)}");
                            break;
                    }
                    fileNameIndex++;
                }
                // save document

                var pdfRenderer = new PdfDocumentRenderer(outputFile.Unicode)
                {
                    Document = document
                };


                pdfRenderer.RenderDocument();
                
                if (!options.UseGivenCredentials)
                    pdfRenderer.PdfDocument.Save(fileName);
                else
                {
                    var domainAndUserName = GetDomainAndUserName(options.UserName);
                    using (Impersonation.LogonUser(domainAndUserName[0], domainAndUserName[1], options.Password, LogonType.NewCredentials))
                    {
                        pdfRenderer.PdfDocument.Save(fileName);
                    }
                }

                byte[] resultAsBytes = null;

                if (options.GetResultAsByteArray)
                {
                    using (MemoryStream stream = new MemoryStream())
                    {
                        pdfRenderer.PdfDocument.Save(stream, false);
                        resultAsBytes = stream.ToArray();
                    }
                }

                return new Output { Success = true, FileName = fileName, ResultAsByteArray = resultAsBytes };
            }
            catch (Exception ex)
            {
                if (options.ThrowErrorOnFailure)
                    throw ex;

                return new Output { Success = false };
            }
        }

        /// <summary>
        /// Define page parameters.
        /// </summary>
        /// <param name="setup"></param>
        /// <param name="pageWidth"></param>
        /// <param name="pageHeight"></param>
        /// <param name="docDefinition"></param>
        private static void SetupPage(PageSetup setup, Unit pageWidth, Unit pageHeight, DocumentDefinition docDefinition)
        {
            setup.Orientation = docDefinition.PageOrientation.ConvertEnum<Orientation>();
            setup.PageHeight = pageHeight;
            setup.PageWidth = pageWidth;
            setup.LeftMargin = new Unit(docDefinition.MarginLeftInCm, UnitType.Centimeter);
            setup.TopMargin = new Unit(docDefinition.MarginTopInCm, UnitType.Centimeter);
            setup.RightMargin = new Unit(docDefinition.MarginRightInCm, UnitType.Centimeter);
            setup.BottomMargin = new Unit(docDefinition.MarginBottomInCm, UnitType.Centimeter);
        }

        /// <summary>
        /// Define font parameters.
        /// </summary>
        /// <param name="style"></param>
        /// <param name="settings"></param>
        private static void SetFont(Style style, StyleSettingsDefinition settings)
        {
            style.Font.Name = settings.FontFamily;
            style.Font.Size = new Unit(settings.FontSizeInPt, UnitType.Point);
            style.Font.Color = Colors.Black;

            switch (settings.FontStyle)
            {
                case FontStyleEnum.Bold:
                    style.Font.Bold = true;
                    break;
                case FontStyleEnum.BoldItalic:
                    style.Font.Bold = true;
                    style.Font.Italic = true;
                    break;
                case FontStyleEnum.Italic:
                    style.Font.Italic = true;
                    break;
                case FontStyleEnum.Underline:
                    style.Font.Underline = Underline.Single;
                    break;
            }
        }

        /// <summary>
        /// Define text paragraph parameters.
        /// </summary>
        /// <param name="style"></param>
        /// <param name="settings"></param>
        /// <param name="isTable"></param>
        private static void SetParagraphStyle(Style style, StyleSettingsDefinition settings, bool isTable)
        {
            style.ParagraphFormat.LineSpacing = new Unit(settings.LineSpacingInPt, UnitType.Point);
            if (!isTable)
            {
                style.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly; 
            }
            style.ParagraphFormat.Alignment = settings.Alignment.ConvertEnum<MigraDoc.DocumentObjectModel.ParagraphAlignment>();
            style.ParagraphFormat.SpaceBefore = new Unit(settings.SpacingBeforeInPt, UnitType.Point);
            style.ParagraphFormat.SpaceAfter = new Unit(settings.SpacingAfterInPt, UnitType.Point);
        }

        /// <summary>
        /// Add text content. Reads one word at a time so that multiple whitespaces are added correctly.
        /// </summary>
        /// <param name="section"></param>
        /// <param name="textContent"></param>
        /// <param name="style"></param>
        private static void AddTextContent(Section section, string textContent, Style style)
        {
            // skip if text content if empty
            if (string.IsNullOrWhiteSpace(textContent))
            {
                return;
            }

            var paragraph = section.AddParagraph();
            paragraph.Style = style.Name;

            //read text line by line
            string line;
            using (var reader = new StringReader(textContent))
            {
                while ((line = reader.ReadLine()) != null)
                {
                    // read text one word at a time, so that multiple whitespaces are added correctly
                    foreach (var word in line.Split(new char[] {' ', '\t' }))
                    {
                        paragraph.AddText(word);
                        paragraph.AddSpace(1);
                    }
                    // add newline
                    paragraph.AddLineBreak();
                }
            }
        }

        /// <summary>
        /// Describes how an image should be added to a page section.
        /// </summary>
        /// <param name="section"></param>
        /// <param name="imageDef"></param>
        /// <param name="pageWidth"></param>
        private static void AddImage(Section section, ImageDefinition imageDef, Unit pageWidth)
        {
            Unit originalImageWidthInches;
            // work around to get image dimensions
            using (System.Drawing.Image userImage = System.Drawing.Image.FromFile(imageDef.ImagePath))
            {
                // get image width in inches
                var imageInches = userImage.Width / userImage.VerticalResolution;
                originalImageWidthInches = new Unit(imageInches, UnitType.Inch);
            }

            // add image
            Image image = section.AddImage(imageDef.ImagePath);

            // Calculate Image size: 
            // if actual image size is larger than PageWidth - margins, set image width as page width - margins
            Unit actualPageContentWidth = new Unit((pageWidth.Inch - section.PageSetup.LeftMargin.Inch - section.PageSetup.RightMargin.Inch), UnitType.Inch);

            if (imageDef.ImageWidthInCm > 0 && imageDef.ImageWidthInCm < actualPageContentWidth.Centimeter)
            {
                image.Width = new Unit(imageDef.ImageWidthInCm, UnitType.Centimeter);
            }
            else if (originalImageWidthInches > actualPageContentWidth)
            {
                image.Width = actualPageContentWidth;
            }

            if (imageDef.LockAspectRatio)
            {
                image.LockAspectRatio = imageDef.LockAspectRatio;
            }
            else if (imageDef.ImageHeightInCm > 0)
            {
                image.Height = new Unit(imageDef.ImageHeightInCm, UnitType.Centimeter);
            }

            if (imageDef.Alignment == AlignmentEnum.Center || imageDef.Alignment == AlignmentEnum.Justify)
            {
                image.Left = ShapePosition.Center;
            }
            else
            {
                image.Left = imageDef.Alignment.ConvertEnum<ShapePosition>();
            }
        }

        /// <summary>
        /// Describes how an image should be added to a table cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="imagePath"></param>
        /// <param name="cellWidth"></param>
        private static void AddImage(Cell cell, string imagePath, double cellWidth)
        {
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Path to header graphics was empty or the file does not exist.");
            }

            Image image = cell.AddImage(imagePath);

            image.Width = new Unit(cellWidth, UnitType.Centimeter);
            image.LockAspectRatio = true;
            image.Top = ShapePosition.Top;
            image.Left = ShapePosition.Left;
        }

        /// <summary>
        /// Adds a table to the page.
        /// </summary>
        /// <param name="section"></param>
        /// <param name="tableDef"></param>
        /// <param name="pageWidth"></param>
        /// <param name="style"></param>
        private static void AddTable(Section section, TableDefinition tableDef, Unit pageWidth, Style style)
        {
            Table table;

            switch (tableDef.TableType)
            {
                case TableTypeEnum.Header:
                    table = section.Headers.Primary.AddTable();
                    break;
                case TableTypeEnum.Footer:
                    table = section.Footers.Primary.AddTable();
                    break;
                default:
                    table = section.AddTable();
                    break;
            }

            Unit tableWidth = new Unit(0, UnitType.Centimeter);
            Unit actualPageContentWidth = new Unit((pageWidth.Centimeter - section.PageSetup.LeftMargin.Centimeter - section.PageSetup.RightMargin.Centimeter), UnitType.Centimeter);

            foreach (var column in tableDef.Columns)
            {
                Unit columnWidth = new Unit(column.WidthInCm, UnitType.Centimeter);
                tableWidth += columnWidth;
                if (tableWidth > actualPageContentWidth)
                {
                    throw new Exception($"Page allows table to be {actualPageContentWidth.Centimeter} cm wide. Provided table's width is larger than that, {tableWidth.Centimeter} cm.");
                }

                table.AddColumn(columnWidth);
            }


            if (tableDef.HasHeaderRow)
            {
                var columnHeaders = tableDef.Columns.Select(column => column.Name).ToList();
                var headerColumnDefinitions = new List<ColumnDefinition>();
                for (int i = 0; i < columnHeaders.Count; i++)
                {
                    headerColumnDefinitions.Add(new ColumnDefinition { Type = ColumnTypeEnum.Text });
                }
                ProcessRow(table, headerColumnDefinitions, columnHeaders, style);
            }

            foreach (var dataRow in tableDef.RowData)
            {
                var data = dataRow.Select(row => row.Value).ToList();
                ProcessRow(table, tableDef.Columns, data, style);
            }

            if (tableDef.StyleSettings.BorderWidthInPt > 0)
            {
                switch (tableDef.StyleSettings.BorderStyle)
                {
                    case BorderStyleEnum.Top:
                        table.Borders.Top.Width = new Unit(tableDef.StyleSettings.BorderWidthInPt, UnitType.Point);
                        break;
                    case BorderStyleEnum.Bottom:
                        table.Borders.Bottom.Width = new Unit(tableDef.StyleSettings.BorderWidthInPt, UnitType.Point);
                        break;
                    case BorderStyleEnum.All:
                        table.Borders.Width = new Unit(tableDef.StyleSettings.BorderWidthInPt, UnitType.Point);
                        break;
                    case BorderStyleEnum.None:
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Helper method to process single rows of a table.
        /// </summary>
        /// <param name="table"></param>
        /// <param name="columns"></param>
        /// <param name="data"></param>
        /// <param name="style"></param>
        private static void ProcessRow(Table table, List<ColumnDefinition> columns, List<string> data, Style style)
        {
            var row = table.AddRow();
            row.VerticalAlignment = VerticalAlignment.Center;

            for (int i = 0; i < data.Count; i++)
            {
                switch (columns[i].Type)
                {
                    case ColumnTypeEnum.Text:
                        var textField = row.Cells[i].AddParagraph();
                        textField.Style = style.Name;
                        textField.AddText(data[i]);
                        break;
                    case ColumnTypeEnum.Image:
                        AddImage(row.Cells[i], data[i], columns[i].WidthInCm);
                        break;
                    case ColumnTypeEnum.PageNum:
                        var pagenumField = row.Cells[i].AddParagraph();
                        pagenumField.Style = style.Name;
                        pagenumField.AddPageField();
                        pagenumField.AddText(" (");
                        pagenumField.AddNumPagesField();
                        pagenumField.AddText(")");
                        break;
                }
            }
        }

        /// <summary>
        /// Helper method to parse domain from a username.
        /// </summary>
        /// <param name="username"></param>
        /// <returns></returns>
        private static string[] GetDomainAndUserName(string username)
        {
            var domainAndUserName = username.Split('\\');
            if (domainAndUserName.Length != 2)
            {
                throw new ArgumentException($@"UserName field must be of format domain\username was: {username}");
            }
            return domainAndUserName;
        }
    }
}
