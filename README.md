# Frends.Community.PdfFromTemplate
Frends wrapper task for PDFSharp library. Task takes in document specification in JSON format. 

- [Installing](#installing)
- [Tasks](#tasks)
  - [Create Pdf](#createpdf)
- [License](#license)
- [Building](#building)
- [Contributing](#contributing)
- [Change Log](#change-log)

# Installing
You can install the task via FRENDS UI Task View or you can find the nuget package from the following nuget feed:
`nuget feed missing`.

# Tasks

## CreatePdf

### Task Properties

### Output File

Settings for writing the PDF file.

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Save to disk | bool | Should the generated document be saved to disk. Output is available as a byte array as well. | true |
| Directory | string | Destination folder for the PDF file created. | C:\Pdf_Output |
| File name | string | File name for the PDF file created. | my_pdf_file.pdf |
| File Exists Action | enum {Error, Overwrite, Rename} | What to do if destination file already exists. Error: throw exception. Overwrite: Replaces existing file. Rename: Renames file by adding '\_(1)' to the end. (pdf_file.pdf --> pdf_file\_(1).pdf) |
| Unicode | bool | If true, Unicode support is enabled. Otherwise ANSI characters are supported. Further [documentation](http://www.pdfsharp.net/wiki/Unicode-sample.ashx). | false |


### Content

Actual PDF document content

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Content JSON | string | JSON definition of the PDF document to produce. | Please see description of template format. |


### Options

Processing options.

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Use Given Credentials | bool | If set, allows you to give the user credentials to use to write the PDF file on remote hosts. | true |
| User Name | string | Domain and username. | 'mydomain\username' |
| Password | string | Password of user. | |
| Throw Error on failure | bool | True: Throws error if PDF writer Task fails. False: Returns object { Success = false } if Task fails. | true |
| Get result as byte array | bool | If set to true, task output will contain resulting document as a byte array. True by default. | true |


### Result

Contents of the resulting object.

| Property             | Type                 | Description                          | Example |
| ---------------------| ---------------------| ------------------------------------ | ----- |
| Success | bool | Task execution result status. | true |
| FileName | string | Full path to PDF document created. | c:\output\example_pdf.pdf |
| ResultAsByteArray | byte[] | Resulting document as byte array | n/a |


# Description of accepted document JSON

A schema file is provided, which describes accepted document format in JSON markup. Please look up `schemas\PdfFromTemplate_schema.json`.

A model document JSON:

```
{
    "PageSize": "A4",
    "PageOrientation": "Portrait",
    "Title": "Document",
    "Author": "FRENDS",
    "MarginLeftInCm": 2.5,
    "MarginTopInCm": 2.5,
    "MarginRightInCm": 2.5,
    "MarginBottomInCm": 2.5,
    "DocumentElements": [
        {
            "HasHeaderRow": true,
            "TableType": "Table",
            "StyleSettings": {
                "FontFamily": "Times New Roman",
                "FontSizeInPt": 10,
                "FontStyle": "Italic",
                "LineSpacingInPt": 0,
                "Alignment": "Left",
                "SpacingBeforeInPt": 0,
                "SpacingAfterInPt": 0,
                "BorderWidthInPt": 0.5,
                "BorderStyle": "All"
            },
            "Columns": [
                {
                    "Name": "Column 1",
                    "WidthInCm": 4,
                    "Type": "Text"
                },
                {
                    "Name": "Column 2",
                    "WidthInCm": 5,
                    "Type": "Text"
                },
                {
                    "Name": "Column 3",
                    "WidthInCm": 6,
                    "Type": "Text"
                }
            ],
            "RowData": [
                {
                    "Column 1": "Cell 1",
                    "Column 2": "Cell 2",
                    "Column 3": "Cell 3"
                },
                {
                    "Column 1": "Cell 4",
                    "Column 2": "Cell 5",
                    "Column 3": "Cell 6"
                }
            ]
        },
        {
            "HasHeaderRow": false,
            "TableType": "Header",
            "StyleSettings": {
                "FontFamily": "Times New Roman",
                "FontSizeInPt": 10,
                "FontStyle": "Bold",
                "LineSpacingInPt": 0,
                "Alignment": "Left",
                "SpacingBeforeInPt": 0,
                "SpacingAfterInPt": 0,
                "BorderWidthInPt": 0.5,
                "BorderStyle": "All"
            },
            "Columns": [
                {
                    "Name": "Column 1",
                    "WidthInCm": 4,
                    "Type": "Text"
                },
                {
                    "Name": "Column 2",
                    "WidthInCm": 5,
                    "Type": "Text"
                },
                {
                    "Name": "Column 3",
                    "WidthInCm": 6,
                    "Type": "Text"
                }
            ],
            "RowData": [
                {
                    "Column 1": "Cell 1",
                    "Column 2": "Cell 2",
                    "Column 3": "Cell 3"
                }
            ]
        },
        {
            "InsertPageBreak": true
        },
        {
            "Text": "Paragraph text can be typed here.\nNew line characters give you new lines.",
            "StyleSettings": {
                "FontFamily": "Times New Roman",
                "FontSizeInPt": 10,
                "FontStyle": "Regular",
                "LineSpacingInPt": 14,
                "Alignment": "Left",
                "SpacingBeforeInPt": 8,
                "SpacingAfterInPt": 0,
                "BorderWidthInPt": 0.5,
                "BorderStyle": "All"
            }
        },
        {
            "ImagePath": "C:\\img\\logo.jpg",
            "Alignment": "Left",
            "LockAspectRatio": false,
            "ImageWidthInCm": 2.5,
            "ImageHeightInCm": 2.5
        }
    ]
}
```

The document object must have the following main-level elements:

* PageSize: (enum) Tells what should be the resulting page size for the document.
  * Allowed options: A0, A1, A2, A3, A4, A5, A6, B5, Ledger, Legal, Letter
* PageOrientation: (enum) Specifies the orientation of the page.
  * Allowed options: Portrait, Landscape
* MarginLeftInCm (number)
* MarginTopInCm (number)
* MarginRightInCm (number)
* MarginBottomInCm (number)
  * Margins of the page, in cm
* DocumentElements: (array) An array containing the actual page elements.
  * Array can also be empty, but what's the point of generating blank pages?

The document may also have the following optional main-level elements:
* Title: (string) Optional title of the document.
  * This will appear in the file preferences and hence it will not be printed on the result.
* Author: (string) Optional author of the document.
  * This will appear in the file preferences and hence it will not be printed on the result.

## DocumentElements' description

### PageBreak

Indicates that between previous and next elements should appear a page break. The element has only one field:

* InsertPageBreak (bool): if set to true, inserts a page break between previous and next elements.

### Paragraph

Element for text content. Has the following fields:

* Text: (string) the actual text content of the element.
* StyleSettings: (StyleSettings element) how the text should be formatted. Please refer to StyleSettings a few rows below for details.

### Table

Describes a table element. A table can be inserted as a page content element, or into header or footer. (There is no support for different footers on different pages; footer is the same for all pages.)

A table element has the following fields:

* HasHeaderRow: (bool) indicates whether the table should have a header row, where all the column names are printed.
* TableType: (enum) tells where the table should be put: to page contents, or to header or footer.
  * Allowed options: Table, Header, Footer
* StyleSettings: (StyleSettings element) how the text should be formatted. Please refer to StyleSettings a few rows below for details.
* Columns: (object) describes the columns of the table. NB. If the sum of columns widths' is larger than (pagesize - (left margin) - (right margin)), an exception will be thrown.
  * Name: (string) name of the column. If HasHeaderRow is set to true, this value is printed on the first row.
  * Type: (enum) type of the column. A column (meaning the entire column) can have three types of information:
    * Text
    * Image: eg. if header/footer should contain some graphics. Then the row's value should contain full path to the desired image file.
    * PageNum: eg. if header/footer should contain page number. Page numbers will be printed in format `No. of this page (No. of total page count)`
  * WidthInCm: (number) column's width in cm.
* RowData: (array of <key, value> objects) Describes the row data of the table. Array consists of object containing key-value-objects, where the keys are column names and values the data to be printed.

### Image

Describes how an image should be printed on the page. PNG graphics are supported, and JPG files seem to work as well.

An image element has the following fields:

* ImagePath: (string) path to the image file to be printed.
  * This is a required field for this element.
* Alignment: (enum) how the image should be aligned on the page.
  * Allowed options: Left, Center, Right
  * Default is center.
* LockAspectRatio: (bool) indicates whether the aspect ratio should be preserved. NB. cannot be used with true value in conjuction with ImageHeightInCm.
* ImageWidthInCm: (number) width of the printed image in cm.
  * If width is set to be larger than (pagesize - (left margin) - (right margin)), width will be set to actual page content size.
* ImageHeightInCm: (number) height of the printed image in cm. Can only be used if LockAspectRatio is set to false.

## StyleSettings

These are applied to two DocumentElement types: Tables and Paragraphs.

* FontFamily: (string) which font should be used.
* FontSizeInPt: (number) font's size in points.
* FontStyle: (enum) style of the font.
  * Allowed options: Regular, Bold, Italic, BoldItalic, Underline
* LineSpacingInPt: (number) space between lines in points
* Alignment: (enum) how the text should be aligned
  * Allowed options: Left, Center, Justify, Right
* SpacingBeforeInPt: (number) space between previous and this element in points
* SpacingAfterInPt: (number) space between this and next element in points
* BorderWidthInPt: (number) only for tables: width of border line in points
* BorderStyle: (enum) only for tables: style of border line
  * Allowed options: None, Top, Bottom, All

# License

This project is licensed under the MIT License - see the LICENSE file for details

# Building

Clone a copy of the repo

`git clone https://github.com/CommunityHiQ/Frends.Community.PdfFromTemplate.git`

Restore dependencies

`dotnet restore`

Rebuild the project

`dotnet build`

Run tests   

`dotnet test`

Create a nuget package

`dotnet pack --configuration Release Frends.Community.PdfFromTemplate`

# Contributing
When contributing to this repository, please first discuss the change you wish to make via issue, email, or any other method with the owners of this repository before making a change.

1. Fork the repo on GitHub
2. Clone the project to your own machine
3. Commit changes to your own branch
4. Push your work back up to your fork
5. Submit a Pull request so that we can review your changes

NOTE: Be sure to merge the latest from "upstream" before making a pull request!

# Change Log

| Version             | Changes                 |
| ---------------------| ---------------------|
| 1.0.0 | Initial version of PdfFromTemplate |
