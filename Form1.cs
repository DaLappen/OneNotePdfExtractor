namespace oneToPdf;
using System;
using System.IO;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml;
using System.Text.RegularExpressions;
using System.Text;
using OneNote = Microsoft.Office.Interop.OneNote;
using Microsoft.VisualBasic.Logging;

public partial class Form1 : Form
{
    private OneNote.Application onApplication;

    public Form1()
    {
        InitializeComponent();
        onApplication = new OneNote.Application();
    }

    private void btnExtractPdfs_Click(object sender, EventArgs e)
    {
        try
        {
            string outputFolder = ShowFolderDialog();
            if (string.IsNullOrEmpty(outputFolder))
                return;

            ExtractPdfsFromOneNote(outputFolder);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private string ShowFolderDialog()
    {
        using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
        {
            folderDialog.Description = "Select output folder for extracted PDFs";
            folderDialog.UseDescriptionForTitle = true;

            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                return folderDialog.SelectedPath;
            }
            return null;
        }
    }

    private void ExtractPdfsFromOneNote(string outputFolder)
    {
        string hierarchyXml;
        onApplication.GetHierarchy(null, OneNote.HierarchyScope.hsPages, out hierarchyXml);

        XDocument doc = XDocument.Parse(hierarchyXml);
        // Define the namespace explicitly instead of using GetDefaultNamespace()
        XNamespace ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        int extractedCount = 0;
        StringBuilder logBuilder = new StringBuilder();
        logBuilder.AppendLine("OneNote PDF Extraction Log:");
        logBuilder.AppendLine("==========================");

        // Find all pages across the hierarchy
        foreach (var page in doc.Descendants(ns + "Page"))
        {
            string pageId = page.Attribute("ID")?.Value;
            string pageName = page.Attribute("name")?.Value ?? "Untitled Page";

            // Find the section and notebook this page belongs to
            var section = page.Parent;
            string sectionName = section?.Attribute("name")?.Value ?? "Unknown Section";

            logBuilder.AppendLine($"Processing page: {pageName} in section: {sectionName}");

            if (string.IsNullOrEmpty(pageId))
            {
                logBuilder.AppendLine("  Skipped: No page ID found");
                continue;
            }

            try
            {
                string pageContentXml;
                onApplication.GetPageContent(pageId, out pageContentXml, OneNote.PageInfo.piAll);

                int pdfCount = ExtractPdfsFromPage(pageContentXml, outputFolder, $"{sectionName}_{pageName}", logBuilder);
                extractedCount += pdfCount;

                logBuilder.AppendLine($"  Found {pdfCount} PDF(s) on this page");
            }
            catch (Exception ex)
            {
                logBuilder.AppendLine($"  Error processing page: {ex.Message}");
            }
        }

        // Display results
        if (extractedCount > 0)
        {
            MessageBox.Show($"Extraction complete! Found {extractedCount} PDF files.\n\nSaved to: {outputFolder}",
                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            MessageBox.Show("No PDF files were found in your OneNote notebooks.",
                "Extraction Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Option to show detailed log
        if (MessageBox.Show("Would you like to see the detailed extraction log?",
            "Extraction Log", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
        {
            ShowLog(logBuilder.ToString());
        }
    }
    private int ExtractPdfsFromPage(string pageXml, string outputFolder, string pageIdentifier, StringBuilder log)
    {
        int count = 0;
        try
        {
            XDocument pageDoc = XDocument.Parse(pageXml);
            // Use explicit namespace
            XNamespace ns = "http://schemas.microsoft.com/office/onenote/2013/onenote";

            // Look for embedded files (OneNote stores them in InsertedFile elements)
            foreach (var insertedFile in pageDoc.Descendants(ns + "InsertedFile"))
            {
                try
                {
                    // Check if it's a PDF
                    var pathAttr = insertedFile.Attribute("pathCache");
                    if (pathAttr == null)
                    {
                        log.AppendLine("    Found inserted file but no path attribute");
                        continue;
                    }

                    string filePath = pathAttr.Value;
                    log.AppendLine($"    Found inserted file: {filePath}");

                    if (!filePath.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
                    {
                        log.AppendLine("      Skipped: Not a BIN file");
                        continue;
                    }

                    // Look for the corresponding .bin file
                    string binFilePath = GetBinFilePath(filePath);
                    log.AppendLine($"      Looking for .bin file at: {binFilePath}");

                    if (!File.Exists(binFilePath))
                    {
                        log.AppendLine("      Error: Could not find corresponding .bin file");
                        continue;
                    }

                    // Generate target PDF filename
                    string originalFilename = Path.GetFileName(filePath);
                    string sanitizedName = string.Join("_", pageIdentifier.Split(Path.GetInvalidFileNameChars()));

                    string fileName = Path.Combine(outputFolder, $"{sanitizedName}_{originalFilename}");

                    // Ensure filename is not too long
                    if (fileName.Length > 260)
                    {
                        // Truncate the path but keep the extension
                        fileName = fileName.Substring(0, 255) + ".pdf";
                        log.AppendLine("      Warning: Filename was truncated due to length");
                    }

                    // If the file exists, add a number
                    string baseFileName = Path.Combine(Path.GetDirectoryName(fileName), Path.GetFileNameWithoutExtension(fileName));
                    string extension = ".pdf";
                    int fileCounter = 1;
                    fileName = $"{baseFileName}{extension}";
                    while (File.Exists(fileName))
                    {
                        fileName = $"{baseFileName}_{fileCounter++}{extension}";
                    }

                    // Copy/Move the bin file to the output location as PDF
                    File.Copy(binFilePath, fileName);
                    log.AppendLine($"      Extracted PDF to: {fileName}");
                    count++;
                }
                catch (Exception ex)
                {
                    log.AppendLine($"      Error extracting file: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            log.AppendLine($"    Error processing page XML: {ex.Message}");
        }

        return count;
    }

    // Helper function to find the bin file corresponding to the PDF
    private string GetBinFilePath(string pdfPath)
    {
        try
        {
            // The .bin files might be stored in several possible locations

            // Option 1: Direct .bin extension replacement
            string binPath1 = Path.ChangeExtension(pdfPath, ".bin");
            if (File.Exists(binPath1))
                return binPath1;

            // Option 2: OneNote cache location
            string oneNoteCachePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft", "OneNote", "16.0", "cache");

            // Sometimes OneNote stores files using hash values or IDs
            // Try to find by filename without extension
            string fileName = Path.GetFileNameWithoutExtension(pdfPath);

            if (Directory.Exists(oneNoteCachePath))
            {
                // Look for any .bin file matching our filename pattern
                var matchingFiles = Directory.GetFiles(oneNoteCachePath, "*.bin")
                    .Where(f => Path.GetFileName(f).Contains(fileName))
                    .ToList();

                if (matchingFiles.Any())
                    return matchingFiles.First();
            }

            // Option 3: Sometimes OneNote uses a temporary location
            string tempPath = Path.Combine(Path.GetTempPath(), "OneNote");
            if (Directory.Exists(tempPath))
            {
                var tempFiles = Directory.GetFiles(tempPath, "*.bin", SearchOption.AllDirectories);
                foreach (var file in tempFiles)
                {
                    // Check if this bin file contains our PDF content
                    if (IsPdfFile(file))
                    {
                        return file;
                    }


                }
            }

            // As a last resort, return the original path with .bin extension
            return binPath1;
        }
        catch
        {
            // If anything fails, just return a simple extension swap
            return Path.ChangeExtension(pdfPath, ".bin");
        }
    }

    // Helper to check if a file is actually a PDF
    private bool IsPdfFile(string filePath)
    {
        try
        {
            if (!File.Exists(filePath))
                return false;

            // Check for PDF signature in first few bytes
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                if (fs.Length < 5)
                    return false;

                byte[] header = new byte[5];
                fs.Read(header, 0, 5);

                // PDF files start with "%PDF-"
                string headerString = System.Text.Encoding.ASCII.GetString(header);
                return headerString.StartsWith("%PDF-");
            }
        }
        catch
        {
            return false;
        }
    }
    // Helper method to show the log in a dialog
    private void ShowLog(string logText)
    {
        Form logForm = new Form
        {
            Text = "Extraction Log",
            Size = new Size(700, 500),
            StartPosition = FormStartPosition.CenterParent
        };

        TextBox textBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            Dock = DockStyle.Fill,
            Text = logText,
            Font = new Font("Consolas", 9)
        };

        Button saveButton = new Button
        {
            Text = "Save Log",
            Dock = DockStyle.Bottom
        };

        saveButton.Click += (sender, e) =>
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                saveDialog.FileName = "OneNotePdfExtraction.log";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(saveDialog.FileName, logText);
                    MessageBox.Show($"Log saved to {saveDialog.FileName}", "Log Saved",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        };

        logForm.Controls.Add(textBox);
        logForm.Controls.Add(saveButton);
        logForm.ShowDialog();
    }
}
