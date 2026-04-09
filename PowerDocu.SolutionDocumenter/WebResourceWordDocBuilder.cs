using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerDocu.Common;

namespace PowerDocu.SolutionDocumenter
{
    class WebResourceWordDocBuilder : WordDocBuilder
    {
        private readonly SolutionDocumentationContent content;

        public WebResourceWordDocBuilder(SolutionDocumentationContent contentDocumentation, string template)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);

            List<WebResourceEntity> webResources = content.solution.Customizations.getWebResources();
            if (webResources.Count == 0) return;

            string filename = InitializeWordDocument(
                content.folderPath + "WebResources - " + content.filename, template);
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true);
            mainPart = wordDocument.MainDocumentPart;
            body = mainPart.Document.Body;
            PrepareDocument(!string.IsNullOrEmpty(template));

            ExtractWebResourceContent(webResources);
            BuildOverview(webResources);
            BuildImageSection(webResources);
            BuildTextResourceSections(webResources);

            NotificationHelper.SendNotification("Created Word documentation for web resources in " + content.solution.UniqueName);
        }

        private void ExtractWebResourceContent(List<WebResourceEntity> webResources)
        {
            string zipPath = content.context?.SourceZipPath;
            if (string.IsNullOrEmpty(zipPath) || !File.Exists(zipPath)) return;

            string wrOutputDir = Path.Combine(content.folderPath, "WebResources");
            Directory.CreateDirectory(wrOutputDir);

            using FileStream zipStream = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Read);

            foreach (WebResourceEntity wr in webResources)
            {
                if (string.IsNullOrEmpty(wr.FileName)) continue;

                string entryPath = wr.FileName.TrimStart('/');
                ZipArchiveEntry entry = archive.Entries.FirstOrDefault(e =>
                    e.FullName.Equals(entryPath, StringComparison.OrdinalIgnoreCase));
                if (entry == null) continue;

                using Stream entryStream = entry.Open();
                using MemoryStream ms = new MemoryStream();
                entryStream.CopyTo(ms);
                wr.Content = ms.ToArray();

                if (wr.IsImageType())
                {
                    string safeFileName = CharsetHelper.GetSafeName(wr.Name) + wr.GetFileExtension();
                    string outputFile = Path.Combine(wrOutputDir, safeFileName);
                    File.WriteAllBytes(outputFile, wr.Content);
                }
            }
        }

        private void BuildOverview(List<WebResourceEntity> webResources)
        {
            AddHeading("Web Resources - " + content.solution.UniqueName, "Heading1");
            body.AppendChild(new Paragraph(new Run(new Text(
                $"This solution contains {webResources.Count} web resource(s)."))));

            AddHeading("Overview", "Heading2");
            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Display Name"), new Text("Name"), new Text("Type"), new Text("Version")));
            foreach (WebResourceEntity wr in webResources.OrderBy(w => w.DisplayName ?? w.Name))
            {
                table.Append(CreateRow(
                    new Text(!string.IsNullOrEmpty(wr.DisplayName) ? wr.DisplayName : wr.Name ?? ""),
                    new Text(wr.Name ?? ""),
                    new Text(wr.GetTypeDisplayName()),
                    new Text(wr.IntroducedVersion ?? "")));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run()));
        }

        private void BuildImageSection(List<WebResourceEntity> webResources)
        {
            var imageResources = webResources
                .Where(w => w.IsImageType() && w.Content != null)
                .OrderBy(w => w.DisplayName ?? w.Name).ToList();
            if (imageResources.Count == 0) return;

            AddHeading("Images", "Heading2");
            foreach (WebResourceEntity wr in imageResources)
            {
                string displayName = wr.DisplayName ?? wr.Name ?? "";
                AddHeading(displayName, "Heading3");

                Table table = CreateTable();
                table.Append(CreateRow(new Text("Name"), new Text(wr.Name ?? "")));
                table.Append(CreateRow(new Text("Type"), new Text(wr.GetTypeDisplayName())));
                table.Append(CreateRow(new Text("Version"), new Text(wr.IntroducedVersion ?? "")));
                table.Append(CreateRow(new Text("File Size"), new Text(FormatFileSize(wr.Content.Length))));
                body.Append(table);
                body.AppendChild(new Paragraph(new Run()));

                InsertWebResourceImage(wr);
            }
        }

        private void BuildTextResourceSections(List<WebResourceEntity> webResources)
        {
            var textResources = webResources
                .Where(w => w.IsTextType() && w.Content != null)
                .OrderBy(w => w.DisplayName ?? w.Name).ToList();
            if (textResources.Count == 0) return;

            AddHeading("Scripts and Markup", "Heading2");
            foreach (WebResourceEntity wr in textResources)
            {
                string displayName = wr.DisplayName ?? wr.Name ?? "";
                AddHeading(displayName, "Heading3");

                Table table = CreateTable();
                table.Append(CreateRow(new Text("Internal Name"), new Text(wr.Name ?? "")));
                table.Append(CreateRow(new Text("Type"), new Text(wr.GetTypeDisplayName())));
                table.Append(CreateRow(new Text("Introduced Version"), new Text(wr.IntroducedVersion ?? "")));
                table.Append(CreateRow(new Text("Is Customizable"), new Text(wr.IsCustomizable ? "Yes" : "No")));
                table.Append(CreateRow(new Text("Is Hidden"), new Text(wr.IsHidden ? "Yes" : "No")));
                table.Append(CreateRow(new Text("File Size"), new Text(FormatFileSize(wr.Content.Length))));
                body.Append(table);
                body.AppendChild(new Paragraph(new Run()));

                string textContent = Encoding.UTF8.GetString(wr.Content);
                if (textContent.Length > 50000)
                {
                    textContent = textContent.Substring(0, 50000) + "\n... (truncated)";
                }
                AddSourceCodeBlock(textContent);

                if (wr.IsImageType())
                {
                    AddHeading("Preview", "Heading4");
                    InsertWebResourceImage(wr);
                }
            }
        }

        private void InsertWebResourceImage(WebResourceEntity wr)
        {
            if (wr.Content == null) return;

            string wrOutputDir = Path.Combine(content.folderPath, "WebResources");
            string safeFileName = CharsetHelper.GetSafeName(wr.Name) + wr.GetFileExtension();
            string imgFile = Path.Combine(wrOutputDir, safeFileName);
            if (!File.Exists(imgFile)) return;

            if (wr.WebResourceType == "11")
            {
                string svgContent = Encoding.UTF8.GetString(wr.Content);
                body.AppendChild(new Paragraph(new Run(
                    InsertSvgImage(mainPart, svgContent, 400, 300))));
                body.AppendChild(new Paragraph(new Run()));
                return;
            }

            try
            {
                int imgWidth, imgHeight;
                using (FileStream stream = new FileStream(imgFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    using (var image = Image.FromStream(stream, false, false))
                    {
                        imgWidth = image.Width;
                        imgHeight = image.Height;
                    }
                    stream.Position = 0;
                    ImagePart imagePart = wr.WebResourceType switch
                    {
                        "6" => mainPart.AddImagePart(ImagePartType.Jpeg),
                        "7" => mainPart.AddImagePart(ImagePartType.Gif),
                        _ => mainPart.AddImagePart(ImagePartType.Png)
                    };
                    imagePart.FeedData(stream);
                    body.AppendChild(new Paragraph(new Run(
                        InsertImage(mainPart.GetIdOfPart(imagePart), imgWidth, imgHeight))));
                }
                body.AppendChild(new Paragraph(new Run()));
            }
            catch (Exception)
            {
                body.AppendChild(new Paragraph(new Run(new Text("[Image could not be embedded]"))));
            }
        }

        private void AddSourceCodeBlock(string code)
        {
            string[] lines = code.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            foreach (string line in lines)
            {
                var run = new Run(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
                run.RunProperties = new RunProperties(
                    new RunFonts() { Ascii = "Consolas", HighAnsi = "Consolas" },
                    new FontSize() { Val = "16" }
                );
                body.AppendChild(new Paragraph(run));
            }
            body.AppendChild(new Paragraph(new Run()));
        }

        private static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("F1") + " KB";
            return (bytes / (1024.0 * 1024.0)).ToString("F1") + " MB";
        }
    }
}
