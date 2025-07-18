using NUnit.Framework;
using System;
using System.IO;
using System.Threading;

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
            try
            {
                _folder = Path.Combine(Path.GetTempPath(), "pdfwriter_tests");
                _destinationFullPath = Path.Combine(_folder, _fileName);

                // Ensure parent directory exists
                var parentDir = Path.GetDirectoryName(_folder);
                if (!Directory.Exists(parentDir))
                {
                    Directory.CreateDirectory(parentDir);
                }

                // Force garbage collection before cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // Temporarily rename output file instead of trying to delete it
                if (Directory.Exists(_folder))
                {
                    foreach (var file in Directory.GetFiles(_folder))
                    {
                        try
                        {
                            // Just grab a unique name without deletion
                            var uniqueName = Path.Combine(
                                Path.GetDirectoryName(file),
                                Path.GetFileNameWithoutExtension(file) + "_" + Guid.NewGuid().ToString().Substring(0, 8) + Path.GetExtension(file)
                            );
                            
                            // Try to rename it if it's locked
                            if (File.Exists(file))
                            {
                                try 
                                {
                                    File.Move(file, uniqueName);
                                }
                                catch
                                {
                                    // Ignore failed move operations
                                    Console.WriteLine($"Could not rename locked file: {file}");
                                }
                            }
                        }
                        catch
                        {
                            // Ignore errors
                        }
                    }
                }
                else
                {
                    Directory.CreateDirectory(_folder);
                }

                _fileProperties = new FileProperties { Directory = _folder, FileName = _fileName, FileExistsAction = FileExistsActionEnum.Error, Unicode = true, SaveToDisk = true };
                _options = new Options { UseGivenCredentials = false, ThrowErrorOnFailure = true, GetResultAsByteArray = true };

                var contentPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"TestFiles\ModelDocument.json");
                var contentDefinition = File.ReadAllText(contentPath);
                _content = new DocumentContent { ContentJson = contentDefinition };
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test setup failed: {ex}");
                throw;
            }
        }

        [TearDown]
        public void TestTearDown()
        {
            try
            {
                if (Directory.Exists(_folder))
                {
                    // Force garbage collection before cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                   
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test teardown warning: {ex}");
                // Don't throw from teardown
            }
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
                + @""", ""Sarake 2"": ""T�m� on keskimm�isen sarakkeen teksti"", ""Sarake 3"": """" } ] }]}";

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
            var contentJson = @"{ ""PageSize"": ""A4"", ""PageOrientation"": ""Portrait"", ""Title"": ""Dokumentti"", ""Author"": ""FRENDS"", ""MarginLeftInCm"": 2.5, ""MarginTopInCm"": 2.5, ""MarginRightInCm"": 2.5, ""MarginBottomInCm"": 2.5, ""DocumentElements"": [ { ""HasHeaderRow"": false, ""TableType"": ""Header"", ""StyleSettings"": { ""FontFamily"": ""Times New Roman"", ""FontSizeInPt"": 10, ""FontStyle"": ""Regular"", ""LineSpacingInPt"": 0, ""Alignment"": ""Left"", ""SpacingBeforeInPt"": 0, ""SpacingAfterInPt"": 0, ""BorderWidthInPt"": 0, ""BorderStyle"": ""None"" }, ""Columns"": [ { ""Name"": ""Sarake 1"", ""WidthInCm"": 2.5, ""Type"": ""Image"" }, { ""Name"": ""Sarake 2"", ""WidthInCm"": 7, ""Type"": ""Text"" }, { ""Name"": ""Sarake 3"", ""WidthInCm"": 2, ""Type"": ""PageNum"" } ], ""RowData"": [ { ""Sarake 1"": """", ""Sarake 2"": ""T�m� on keskimm�isen sarakkeen teksti"", ""Sarake 3"": """" } ] }]}";

            var content = new DocumentContent { ContentJson = contentJson };

            var result = Assert.Throws<FileNotFoundException>(() => PdfTask.CreatePdf(_fileProperties, content, _options));
            Assert.IsFalse(File.Exists(_destinationFullPath));
        }

        [Test]
        public void WritePdf_TableWidthTooLarge()
        {
            var contentJson = @"{ ""PageSize"": ""A4"", ""PageOrientation"": ""Portrait"", ""Title"": ""Dokumentti"", ""Author"": ""FRENDS"", ""MarginLeftInCm"": 2.5, ""MarginTopInCm"": 2.5, ""MarginRightInCm"": 2.5, ""MarginBottomInCm"": 2.5, ""DocumentElements"": [ { ""HasHeaderRow"": false, ""TableType"": ""Header"", ""StyleSettings"": { ""FontFamily"": ""Times New Roman"", ""FontSizeInPt"": 10, ""FontStyle"": ""Regular"", ""LineSpacingInPt"": 0, ""Alignment"": ""Left"", ""SpacingBeforeInPt"": 0, ""SpacingAfterInPt"": 0, ""BorderWidthInPt"": 0, ""BorderStyle"": ""None"" }, ""Columns"": [ { ""Name"": ""Sarake 1"", ""WidthInCm"": 22, ""Type"": ""Text"" }, { ""Name"": ""Sarake 2"", ""WidthInCm"": 7, ""Type"": ""Text"" }, { ""Name"": ""Sarake 3"", ""WidthInCm"": 2, ""Type"": ""PageNum"" } ], ""RowData"": [ { ""Sarake 1"": ""Jotain teksti�"", ""Sarake 2"": ""T�m� on keskimm�isen sarakkeen teksti"", ""Sarake 3"": """" } ] }]}";

            var content = new DocumentContent { ContentJson = contentJson };

            var result = Assert.Throws<Exception>(() => PdfTask.CreatePdf(_fileProperties, content, _options));
            Assert.IsFalse(File.Exists(_destinationFullPath));
        }

        [Test]
        public void WritePdf_SimpleLogoOnly()
        {
            try
            {
                // Set a specific name for this test's output file
                string logoTestFileName = "simple_logo_test.pdf";
                _fileProperties.FileName = logoTestFileName;
                _destinationFullPath = Path.Combine(_folder, logoTestFileName);

                var logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"TestFiles\logo.png").Replace("\\", "\\\\");
                Console.WriteLine($"Logo path: {logoPath}");
                Console.WriteLine($"Logo file exists: {File.Exists(logoPath.Replace("\\\\", "\\"))}");

                // Create a simple document with just the logo image
                var contentJson = @"{
                    ""PageSize"": ""A4"",
                    ""PageOrientation"": ""Portrait"",
                    ""Title"": ""Simple Logo Test"",
                    ""Author"": ""Frends Test"",
                    ""MarginLeftInCm"": 2.5,
                    ""MarginTopInCm"": 2.5,
                    ""MarginRightInCm"": 2.5,
                    ""MarginBottomInCm"": 2.5,
                    ""DocumentElements"": [
                        {
                            ""Text"": ""This document contains a logo image below:"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Arial"",
                                ""FontSizeInPt"": 14,
                                ""FontStyle"": ""Bold"",
                                ""LineSpacingInPt"": 16,
                                ""HorizontalAlignment"": ""Center"",
                                ""VerticalAlignment"": ""Center"",
                                ""SpacingBeforeInPt"": 10,
                                ""SpacingAfterInPt"": 20,
                                ""BorderWidthInPt"": 0,
                                ""BorderStyle"": ""None""
                            }
                        },
                        {
                            ""ImagePath"": """ + logoPath + @""",
                            ""Alignment"": ""Center"",
                            ""LockAspectRatio"": true,
                            ""ImageWidthInCm"": 10,
                            ""ImageHeightInCm"": 0
                        }
                    ]
                }";

                var content = new DocumentContent { ContentJson = contentJson };
                
                Console.WriteLine("Creating PDF with simple logo...");
                var result = PdfTask.CreatePdf(_fileProperties, content, _options);
                
                Console.WriteLine($"PDF with simple logo generated at: {_destinationFullPath}");
                Console.WriteLine($"File exists: {File.Exists(_destinationFullPath)}");
                
                Assert.IsTrue(File.Exists(_destinationFullPath));
                Assert.IsTrue(result.Success);
                
                // Pause to allow manual inspection
                Console.WriteLine("Pausing for 10 seconds to allow manual file inspection");
                Thread.Sleep(10000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in WritePdf_SimpleLogoOnly: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                    Console.WriteLine($"Inner stack trace: {ex.InnerException.StackTrace}");
                }
                throw;
            }
        }



        [Test]
        public void RealContent()
        {
            try
            {
                // Set a specific name for this test's output file
                string logoTestFileName = "RealContent.pdf";
                _fileProperties.FileName = logoTestFileName;
                _destinationFullPath = Path.Combine(_folder, logoTestFileName);

                var logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"TestFiles\logo.png").Replace("\\", "\\\\");
                Console.WriteLine($"Logo path: {logoPath}");
                Console.WriteLine($"Logo file exists: {File.Exists(logoPath.Replace("\\\\", "\\"))}");

                // Create a simple document with just the logo image
                var contentJson = @"{
                    ""PageSize"": ""A4"",
                    ""PageOrientation"": ""Portrait"",
                    ""Title"": ""Testi"",
                    ""MarginLeftInCm"": 2.5,
                    ""MarginTopInCm"": 2.5,
                    ""MarginRightInCm"": 2.5,
                    ""MarginBottomInCm"": 2.5,
                    ""DocumentElements"": [
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Pöytäkirjatyyppi"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""Jakokaappi"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Projekti"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""I1111515"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""Text"": ""\nURAKOITSIJATIEDOT\n"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Arial"",
                                ""FontSizeInPt"": 12,
                                ""FontStyle"": ""Bold"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            }
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Urakoitsija"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""eBus"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Tarkastuksen suorittaja"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""Jesse Ryhänen"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Päivämäärä"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""2024-10-29"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""Text"": ""\nTARKASTUSKOHDE\n"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Arial"",
                                ""FontSizeInPt"": 12,
                                ""FontStyle"": ""Bold"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            }
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Projekti"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""I1111515"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Työtilaus"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""Virhestesti 1"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Jakokaapin tunnus"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""J222463"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Jakokaapin nimi/osoite"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""Poikkipuoliaisentie 110"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Kohteen koordinaatit"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""lat: 60,643483490172734 lng:23,75935982912779"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Kohteen työlaji"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""Jakokaappi OTT"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Rakennetun verkon lupien, sopimusten ja suunnitelman mukaisuus"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""On tehty hyväksytyn suunnitelman mukaan\n"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Rakennettu niin ettei ole tulva-alueella"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""OK\n"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Liittyminen vanhaan verkkoon"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""Ei"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""Text"": ""\nOMAN TYÖN TARKASTUS\n"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Arial"",
                                ""FontSizeInPt"": 12,
                                ""FontStyle"": ""Bold"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            }
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Ulkopuolinen rakenne ja merkinnät"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""OK\nJesse testaa\nKuuluu asennukseen"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Jakokaapin maanrakennustyöt"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""OK\n\nKuuluu asennukseen"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Perustus, sen alue ja syvyys jakokaapin Carunan perustusohjeen mukainen"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""OK\n\nKuuluu asennukseen"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Maadoitukset tehty suunnitelman mukaan"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""OK\n\nKuuluu asennukseen"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Kytkimet"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""OK\n\nKuuluu asennukseen"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""PJ-kaapelit ja sulakkeet"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""OK\n\nKuuluu asennukseen"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        },
                        {
                            ""HasHeaderRow"": true,
                            ""TableType"": ""Table"",
                            ""StyleSettings"": {
                                ""FontFamily"": ""Times New Roman"",
                                ""FontSizeInPt"": 10,
                                ""FontStyle"": ""Regular"",
                                ""LineSpacingInPt"": 14,
                                ""Alignment"": ""Left"",
                                ""SpacingBeforeInPt"": 8,
                                ""SpacingAfterInPt"": 0,
                                ""BorderWidthInPt"": 0.5,
                                ""BorderStyle"": ""All""
                            },
                            ""Columns"": [
                                {
                                    ""Name"": ""Kaapeleiden merkinnät ja verkkotietojärjestelmä"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                },
                                {
                                    ""Name"": ""OK\n\nKuuluu asennukseen"",
                                    ""WidthInCm"": 8,
                                    ""Type"": ""Text""
                                }
                            ],
                            ""RowData"": []
                        }
                    ]
                }";

                var content = new DocumentContent { ContentJson = contentJson };
                
                Console.WriteLine("Creating PDF with simple logo...");
                var result = PdfTask.CreatePdf(_fileProperties, content, _options);
                
                Console.WriteLine($"PDF with simple logo generated at: {_destinationFullPath}");
                Console.WriteLine($"File exists: {File.Exists(_destinationFullPath)}");
                
                Assert.IsTrue(File.Exists(_destinationFullPath));
                Assert.IsTrue(result.Success);
                
                // Pause to allow manual inspection
                Console.WriteLine("Pausing for 10 seconds to allow manual file inspection");
                Thread.Sleep(10000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in WritePdf_SimpleLogoOnly: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                    Console.WriteLine($"Inner stack trace: {ex.InnerException.StackTrace}");
                }
                throw;
            }
        }
    }
}