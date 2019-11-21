using NUnit.Framework;
using System;
using System.IO;

namespace Frends.Community.PdfFromTemplate
{
    [TestFixture]
    public class WriterTests
    {
        private FileProperties _fileProperties;
        private DocumentContent _content;
        private Options _options;

        private string _destinationFullPath;
        private readonly string _fileName = "test_output.pdf";
        private string _folder;

        [SetUp]
        public void TestSetup()
        {
            _folder = Path.Combine(Path.GetTempPath(), "pdfwriter_tests");
            _destinationFullPath = Path.Combine(_folder, _fileName);

            if (!Directory.Exists(_folder))
            {
                Directory.CreateDirectory(_folder);
            }

            _fileProperties = new FileProperties { Directory = _folder, FileName = _fileName, FileExistsAction = FileExistsActionEnum.Error, Unicode = true };
            _options = new Options { UseGivenCredentials = false, ThrowErrorOnFailure = true, GetResultAsByteArray = true };

            var contentPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"TestFiles\ModelDocument.json");
            var contentDefinition = File.ReadAllText(contentPath);
            _content = new DocumentContent { ContentJson = contentDefinition };
        }

        [TearDown]
        public void TestTearDown()
        {
            Directory.Delete(_folder, true);
        }

        [Test]
        public void WritePdf()
        {
            _fileProperties.FileExistsAction = FileExistsActionEnum.Overwrite;

            var result = PdfTask.CreatePdf(_fileProperties, _content, _options);

            Assert.IsTrue(File.Exists(_destinationFullPath));
            Assert.IsTrue(result.Success);
            Assert.IsNotNull(result.ResultAsByteArray);
        }

        [Test]
        public void WritePdf_DoesNotFailIfContentIsEmpty()
        {
            var contentJson = @"{ ""PageSize"": ""A4"", ""PageOrientation"": ""Portrait"", ""Title"": ""Dokumentti"", ""Author"": ""FRENDS"", ""MarginLeftInCm"": 2.5, ""MarginTopInCm"": 2.5, ""MarginRightInCm"": 2.5, ""MarginBottomInCm"": 2.5, ""DocumentElements"": [ ] }";

            var content = new DocumentContent { ContentJson = contentJson };

            var result = PdfTask.CreatePdf(_fileProperties, content, _options);
            Assert.IsTrue(File.Exists(_destinationFullPath));
            Assert.IsTrue(result.Success);
            Assert.IsNotNull(result.ResultAsByteArray);
        }


        [Test]
        public void WritePdf_HeaderGraphics()
        {
            var logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"TestFiles\logo.png").Replace("\\", "\\\\");

            var contentJson = @"{ ""PageSize"": ""A4"", ""PageOrientation"": ""Portrait"", ""Title"": ""Dokumentti"", ""Author"": ""FRENDS"", ""MarginLeftInCm"": 2.5, ""MarginTopInCm"": 2.5, ""MarginRightInCm"": 2.5, ""MarginBottomInCm"": 2.5, ""DocumentElements"": [ { ""HasHeaderRow"": false, ""TableType"": ""Header"", ""StyleSettings"": { ""FontFamily"": ""Times New Roman"", ""FontSizeInPt"": 10, ""FontStyle"": ""Regular"", ""LineSpacingInPt"": 0, ""Alignment"": ""Left"", ""SpacingBeforeInPt"": 0, ""SpacingAfterInPt"": 0, ""BorderWidthInPt"": 0, ""BorderStyle"": ""None"" }, ""Columns"": [ { ""Name"": ""Sarake 1"", ""WidthInCm"": 2.5, ""Type"": ""Image"" }, { ""Name"": ""Sarake 2"", ""WidthInCm"": 7, ""Type"": ""Text"" }, { ""Name"": ""Sarake 3"", ""WidthInCm"": 2, ""Type"": ""PageNum"" } ], ""RowData"": [ { ""Sarake 1"": """
                + logoPath
                + @""", ""Sarake 2"": ""Tämä on keskimmäisen sarakkeen teksti"", ""Sarake 3"": """" } ] }]}";

            var content = new DocumentContent { ContentJson = contentJson };

            var result = PdfTask.CreatePdf(_fileProperties, content, _options);
            Assert.IsTrue(File.Exists(_destinationFullPath));
            Assert.IsTrue(result.Success);
            Assert.IsNotNull(result.ResultAsByteArray);
        }

        [Test]
        public void WritePdf_ThrowsExceptionIfFileExists()
        {
            _options.ThrowErrorOnFailure = true;
            _fileProperties.FileExistsAction = FileExistsActionEnum.Error;

            // Run once so file exists
            PdfTask.CreatePdf(_fileProperties, _content, _options);

            var result = Assert.Throws<Exception>(() => PdfTask.CreatePdf(_fileProperties, _content, _options));
        }

        [Test]
        public void WritePdf_RenamesFilesIfAlreadyExists()
        {
            _fileProperties.FileExistsAction = FileExistsActionEnum.Rename;

            // Create three PDFs
            var result1 = PdfTask.CreatePdf(_fileProperties, _content, _options);
            var result2 = PdfTask.CreatePdf(_fileProperties, _content, _options);
            var result3 = PdfTask.CreatePdf(_fileProperties, _content, _options);

            Assert.IsTrue(File.Exists(result1.FileName));
            Assert.AreEqual(_destinationFullPath, result1.FileName);
            Assert.IsTrue(File.Exists(result2.FileName));
            Assert.IsTrue(result2.FileName.Contains("_(1)"));
            Assert.IsTrue(File.Exists(result3.FileName));
            Assert.IsTrue(result3.FileName.Contains("_(2)"));
        }

        [Test]
        public void WritePdf_ImageNotFound()
        {
            var contentJson = @"{ ""PageSize"": ""A4"", ""PageOrientation"": ""Portrait"", ""Title"": ""Dokumentti"", ""Author"": ""FRENDS"", ""MarginLeftInCm"": 2.5, ""MarginTopInCm"": 2.5, ""MarginRightInCm"": 2.5, ""MarginBottomInCm"": 2.5, ""DocumentElements"": [ { ""ImagePath"": ""C:\\img\\logo.jpg"", ""Alignment"": ""Left"", ""LockAspectRatio"": false, ""ImageWidthInCm"": 2.5, ""ImageHeightInCm"": 2.5 } ]}";

            var content = new DocumentContent { ContentJson = contentJson };

            var result = Assert.Throws<FileNotFoundException>(() => PdfTask.CreatePdf(_fileProperties, content, _options));
            Assert.IsFalse(File.Exists(_destinationFullPath));
        }

        [Test]
        public void WritePdf_LogoNotFound()
        {
            var contentJson = @"{ ""PageSize"": ""A4"", ""PageOrientation"": ""Portrait"", ""Title"": ""Dokumentti"", ""Author"": ""FRENDS"", ""MarginLeftInCm"": 2.5, ""MarginTopInCm"": 2.5, ""MarginRightInCm"": 2.5, ""MarginBottomInCm"": 2.5, ""DocumentElements"": [ { ""HasHeaderRow"": false, ""TableType"": ""Header"", ""StyleSettings"": { ""FontFamily"": ""Times New Roman"", ""FontSizeInPt"": 10, ""FontStyle"": ""Regular"", ""LineSpacingInPt"": 0, ""Alignment"": ""Left"", ""SpacingBeforeInPt"": 0, ""SpacingAfterInPt"": 0, ""BorderWidthInPt"": 0, ""BorderStyle"": ""None"" }, ""Columns"": [ { ""Name"": ""Sarake 1"", ""WidthInCm"": 2.5, ""Type"": ""Image"" }, { ""Name"": ""Sarake 2"", ""WidthInCm"": 7, ""Type"": ""Text"" }, { ""Name"": ""Sarake 3"", ""WidthInCm"": 2, ""Type"": ""PageNum"" } ], ""RowData"": [ { ""Sarake 1"": """", ""Sarake 2"": ""Tämä on keskimmäisen sarakkeen teksti"", ""Sarake 3"": """" } ] }]}";

            var content = new DocumentContent { ContentJson = contentJson };

            var result = Assert.Throws<FileNotFoundException>(() => PdfTask.CreatePdf(_fileProperties, content, _options));
            Assert.IsFalse(File.Exists(_destinationFullPath));
        }

        [Test]
        public void WritePdf_TableWidthTooLarge()
        {
            var contentJson = @"{ ""PageSize"": ""A4"", ""PageOrientation"": ""Portrait"", ""Title"": ""Dokumentti"", ""Author"": ""FRENDS"", ""MarginLeftInCm"": 2.5, ""MarginTopInCm"": 2.5, ""MarginRightInCm"": 2.5, ""MarginBottomInCm"": 2.5, ""DocumentElements"": [ { ""HasHeaderRow"": false, ""TableType"": ""Header"", ""StyleSettings"": { ""FontFamily"": ""Times New Roman"", ""FontSizeInPt"": 10, ""FontStyle"": ""Regular"", ""LineSpacingInPt"": 0, ""Alignment"": ""Left"", ""SpacingBeforeInPt"": 0, ""SpacingAfterInPt"": 0, ""BorderWidthInPt"": 0, ""BorderStyle"": ""None"" }, ""Columns"": [ { ""Name"": ""Sarake 1"", ""WidthInCm"": 22, ""Type"": ""Text"" }, { ""Name"": ""Sarake 2"", ""WidthInCm"": 7, ""Type"": ""Text"" }, { ""Name"": ""Sarake 3"", ""WidthInCm"": 2, ""Type"": ""PageNum"" } ], ""RowData"": [ { ""Sarake 1"": ""Jotain tekstiä"", ""Sarake 2"": ""Tämä on keskimmäisen sarakkeen teksti"", ""Sarake 3"": """" } ] }]}";

            var content = new DocumentContent { ContentJson = contentJson };

            var result = Assert.Throws<Exception>(() => PdfTask.CreatePdf(_fileProperties, content, _options));
            Assert.IsFalse(File.Exists(_destinationFullPath));
        }
    }
}