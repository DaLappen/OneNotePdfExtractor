# OneNotePdfExtractor

OneNotePdfExtractor is a Windows application that extracts PDF files embedded within Microsoft OneNote notebooks.

## Description

Many users embed PDF documents within OneNote for reference purposes, but extracting these files can be challenging through the standard OneNote interface. OneNotePdfExtractor addresses this problem by providing a simple tool that:

1. Connects to the Microsoft OneNote application
2. Reads the entire notebook hierarchy
3. Identifies and extracts all embedded PDF files
4. Saves them to a user-specified location with organized naming

## Features

- Extract all PDF files from all notebooks, sections, and pages
- Preserve original filenames when possible
- Create organized filenames based on section and page names
- Generate detailed extraction logs
- Simple, one-click interface

## Architecture

OneNotePdfExtractor uses the OneNote Interop API to interact with OneNote. The application:

1. Gets the entire OneNote hierarchy (notebooks, sections, pages) as XML
2. Parses this XML to find all pages
3. For each page, retrieves detailed page content
4. Locates embedded PDF file references and their cached paths
5. Finds the corresponding binary files on disk
6. Copies these files to the target location with appropriate naming

## Requirements

- Windows operating system
- Microsoft OneNote installed
- .NET Framework 4.7.2 or later

## Installation

1. Clone or download the repository
2. Open the solution in Visual Studio
3. Restore NuGet packages if prompted
4. Build the solution

## Usage

1. Run the application
2. Click the "Extract PDFs from OneNote" button
3. Select an output folder when prompted
4. Wait for the extraction process to complete

## Dependencies

- Microsoft.Office.Interop.OneNote
- System.Xml.Linq
- Windows Forms

## Known Limitations

- OneNote must be installed and accessible on the system
- The tool only extracts PDF files, not other embedded file types
- Very long filenames may be truncated to fit Windows path length limitations

## Troubleshooting

If the tool fails to extract PDFs:

1. Ensure OneNote is installed and can be opened manually
2. Check that you have read access to OneNote cache folders
3. Try running the application as administrator
4. Examine the detailed log for specific error messages

## License

[MIT License](LICENSE)

## Contribution
This project is not actively maintained. 
Contributions, issues, and feature requests are welcome! 