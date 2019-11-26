using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using JsonSubTypes;
#pragma warning disable 1591

namespace Frends.Community.PdfFromTemplate
{
    public enum FileExistsActionEnum { Error, Overwrite, Rename };

    public class FileProperties
    {
        /// <summary>
        /// Should the generated document be saved to disk. Output is available as a byte array as well.
        /// </summary>
        [DefaultValue(true)]
        public bool SaveToDisk { get; set; }
        
        /// <summary>
        /// PDF document destination Directory
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        [DefaultValue(@"C:\Output")]
        public string Directory { get; set; }

        /// <summary>
        /// Filename for created PDF file
        /// </summary>
        [DisplayFormat(DataFormatString = "Text")]
        [DefaultValue("example_file.pdf")]
        public string FileName { get; set; }

        /// <summary>
        /// What to do if destination file already exists
        /// </summary>
        [DefaultValue(FileExistsActionEnum.Error)]
        public FileExistsActionEnum FileExistsAction { get; set; }

        /// <summary>
        /// Use Unicode text (true) or ANSI (false).
        /// </summary>
        [DefaultValue(true)]
        public bool Unicode { get; set; }
    }

    public class DocumentContent
    {
        /// <summary>
        /// JSON definition of the PDF document to produce.
        /// </summary>
        [DisplayFormat(DataFormatString = "Json")]
        public string ContentJson { get; set; }
    }

    public class Options
    {
        /// <summary>
        /// If set, allows you to give the user credentials to use to write the PDF file on remote hosts.
        /// If not set, the agent service user credentials will be used.
        /// </summary>
        [DefaultValue(false)]
        public bool UseGivenCredentials { get; set; }

        /// <summary>
        /// This needs to be of format domain\username
        /// </summary>
        [UIHint(nameof(UseGivenCredentials), "", true)]
        [DisplayFormat(DataFormatString = "Text")]
        [DefaultValue(@"domain\username")]
        public string UserName { get; set; }

        [PasswordPropertyText]
        [UIHint(nameof(UseGivenCredentials), "", true)]
        public string Password { get; set; }

        /// <summary>
        /// True: Throws error on failure
        /// False: Returns object{ Success = false }
        /// </summary>
        [DefaultValue(true)]
        public bool ThrowErrorOnFailure { get; set; }

        /// <summary>
        /// True: Output object contains resulting document as a byte array.
        /// False: Said array is set to null.
        /// </summary>
        [DefaultValue(true)]
        public bool GetResultAsByteArray { get; set; }
    }

    public class Output
    {
        public bool Success { get; set; }

        public string FileName { get; set; }

        public byte[] ResultAsByteArray { get; set; }

        public string ErrorMessage { get; set; }
    }


    // Following classes are for deserialization of JSON input.
    public enum HorizontalAlignmentEnum { Left, Center, Justify, Right };
    public enum VerticalAlignmentEnum { Top, Center, Bottom };
    public enum FontStyleEnum { Regular, Bold, Italic, BoldItalic, Underline };
    public enum TableTypeEnum { Table, Header, Footer };
    public enum ColumnTypeEnum { Text, Image, PageNum };
    public enum BorderStyleEnum { None, Top, Bottom, All };
    public enum PageSizeEnum { A0, A1, A2, A3, A4, A5, A6, B5, Ledger, Legal, Letter };
    public enum PageOrientationEnum { Portrait, Landscape };
    public enum ElementTypeEnum { Paragraph, Image, Table, PageBreak };

    public class ColumnDefinition
    {
        public string Name { get; set; }
        public double WidthInCm { get; set; }
        public ColumnTypeEnum Type { get; set; }
    }

    public class StyleSettingsDefinition
    {
        
        public string FontFamily { get; set; }
        public double FontSizeInPt { get; set; }
        public FontStyleEnum FontStyle { get; set; }
        public double LineSpacingInPt { get; set; }
        public HorizontalAlignmentEnum HorizontalAlignment { get; set; }
        public VerticalAlignmentEnum VerticalAlignment { get; set; }
        public double SpacingBeforeInPt { get; set; }
        public double SpacingAfterInPt { get; set; }
        public double BorderWidthInPt { get; set; }
        public BorderStyleEnum BorderStyle { get; set; }
    }

    [JsonConverter(typeof(JsonSubtypes))]
    [JsonSubtypes.KnownSubTypeWithProperty(typeof(TableDefinition), "TableType")]
    [JsonSubtypes.KnownSubTypeWithProperty(typeof(ParagraphDefinition), "Text")]
    [JsonSubtypes.KnownSubTypeWithProperty(typeof(ImageDefinition), "ImagePath")]
    [JsonSubtypes.KnownSubTypeWithProperty(typeof(PageBreakDefinition), "InsertPageBreak")]
    public abstract class DocumentElement {
        public abstract ElementTypeEnum ElementType { get; }
    }


    public class TableDefinition : DocumentElement
    {
        public override ElementTypeEnum ElementType
        {
            get
            {
                return ElementTypeEnum.Table;
            }
        }
        public bool HasHeaderRow { get; set; }
        public TableTypeEnum TableType { get; set; }
        public StyleSettingsDefinition StyleSettings { get; set; }
        public List<ColumnDefinition> Columns { get; set; }
        public List<Dictionary<string, string>> RowData { get; set; }
    }
    public class ParagraphDefinition : DocumentElement
    {
        public override ElementTypeEnum ElementType
        {
            get
            {
                return ElementTypeEnum.Paragraph;
            }
        }
        public string Text { get; set; }
        public StyleSettingsDefinition StyleSettings { get; set; }
    }
    public class ImageDefinition : DocumentElement
    {
        public override ElementTypeEnum ElementType
        {
            get
            {
                return ElementTypeEnum.Image;
            }
        }
        public string ImagePath { get; set; }
        public HorizontalAlignmentEnum Alignment { get; set; }
        public bool LockAspectRatio { get; set; }
        public double ImageWidthInCm { get; set; }
        public double ImageHeightInCm { get; set; }
    }
    public class PageBreakDefinition : DocumentElement
    {
        public override ElementTypeEnum ElementType
        {
            get
            {
                return ElementTypeEnum.PageBreak;
            }
        }
        public bool InsertPageBreak { get; set; }
    }
    public class DocumentDefinition
    {
        public PageSizeEnum PageSize { get; set; }
        public PageOrientationEnum PageOrientation { get; set; }
        public string Title { get; set; }
        public string Author { get; set; }
        public double MarginLeftInCm { get; set; }
        public double MarginTopInCm { get; set; }
        public double MarginRightInCm { get; set; }
        public double MarginBottomInCm { get; set; }
        public List<DocumentElement> DocumentElements { get; set; }
    }
}
