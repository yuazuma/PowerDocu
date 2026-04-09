using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using PowerDocu.Common;
using Grynwald.MarkdownGenerator;

namespace PowerDocu.SolutionDocumenter
{
    class WebResourceMarkdownBuilder : MarkdownBuilder
    {
        private readonly SolutionDocumentationContent content;
        private readonly string indexFileName;

        public WebResourceMarkdownBuilder(SolutionDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);

            indexFileName = CrossDocLinkHelper.GetWebResourceDocMdPath(content.solution.UniqueName);

            List<WebResourceEntity> webResources = content.solution.Customizations.getWebResources();
            if (webResources.Count == 0) return;

            ExtractWebResourceContent(webResources);
            BuildIndexPage(webResources);
            BuildDetailPages(webResources);

            NotificationHelper.SendNotification("Created Markdown documentation for web resources in " + content.solution.UniqueName);
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

        private void BuildIndexPage(List<WebResourceEntity> webResources)
        {
            MdDocument doc = new MdDocument();
            doc.Root.Add(new MdHeading("Web Resources", 1));
            doc.Root.Add(new MdParagraph($"This solution contains {webResources.Count} web resource(s)."));

            // Overview table
            doc.Root.Add(new MdHeading("Overview", 2));
            List<MdTableRow> rows = new List<MdTableRow>();
            foreach (WebResourceEntity wr in webResources.OrderBy(w => w.DisplayName ?? w.Name))
            {
                string displayName = wr.DisplayName ?? wr.Name ?? "";
                if (wr.IsTextType() && wr.Content != null)
                {
                    string detailPath = CrossDocLinkHelper.GetWebResourceDetailMdPath(content.solution.UniqueName, wr.Name);
                    var nameLink = new MdLinkSpan(displayName, detailPath);
                    rows.Add(new MdTableRow(nameLink, new MdTextSpan(wr.Name ?? ""), new MdTextSpan(wr.GetTypeDisplayName()), new MdTextSpan(wr.IntroducedVersion ?? "")));
                }
                else
                {
                    rows.Add(new MdTableRow(displayName, wr.Name ?? "", wr.GetTypeDisplayName(), wr.IntroducedVersion ?? ""));
                }
            }
            doc.Root.Add(new MdTable(new MdTableRow("Display Name", "Name", "Type", "Version"), rows));

            // Images section
            var imageResources = webResources.Where(w => w.IsImageType() && w.Content != null).OrderBy(w => w.DisplayName ?? w.Name).ToList();
            if (imageResources.Count > 0)
            {
                doc.Root.Add(new MdHeading("Images", 2));
                foreach (WebResourceEntity wr in imageResources)
                {
                    string displayName = wr.DisplayName ?? wr.Name ?? "";
                    doc.Root.Add(new MdHeading(displayName, 3));

                    List<MdTableRow> propRows = new List<MdTableRow>();
                    propRows.Add(new MdTableRow("Name", wr.Name ?? ""));
                    propRows.Add(new MdTableRow("Type", wr.GetTypeDisplayName()));
                    propRows.Add(new MdTableRow("Version", wr.IntroducedVersion ?? ""));
                    propRows.Add(new MdTableRow("File Size", FormatFileSize(wr.Content.Length)));
                    doc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), propRows));

                    string safeFileName = CharsetHelper.GetSafeName(wr.Name) + wr.GetFileExtension();
                    string imgPath = "WebResources/" + safeFileName;
                    doc.Root.Add(new MdParagraph(new MdImageSpan(displayName, imgPath)));
                }
            }

            // Text-based resources section
            var textResources = webResources.Where(w => w.IsTextType() && w.Content != null && !w.IsImageType()).OrderBy(w => w.DisplayName ?? w.Name).ToList();
            if (textResources.Count > 0)
            {
                doc.Root.Add(new MdHeading("Scripts and Markup", 2));
                List<MdTableRow> textRows = new List<MdTableRow>();
                foreach (WebResourceEntity wr in textResources)
                {
                    string displayName = wr.DisplayName ?? wr.Name ?? "";
                    string detailPath = CrossDocLinkHelper.GetWebResourceDetailMdPath(content.solution.UniqueName, wr.Name);
                    var nameLink = new MdLinkSpan(displayName, detailPath);
                    string size = wr.Content != null ? FormatFileSize(wr.Content.Length) : "";
                    textRows.Add(new MdTableRow(nameLink, new MdTextSpan(wr.GetTypeDisplayName()), new MdTextSpan(size)));
                }
                doc.Root.Add(new MdTable(new MdTableRow("Name", "Type", "Size"), textRows));
            }

            doc.Save(Path.Combine(content.folderPath, indexFileName));
        }

        private void BuildDetailPages(List<WebResourceEntity> webResources)
        {
            string wrOutputDir = Path.Combine(content.folderPath, "WebResources");
            Directory.CreateDirectory(wrOutputDir);

            foreach (WebResourceEntity wr in webResources.Where(w => w.IsTextType() && w.Content != null))
            {
                string detailRelPath = CrossDocLinkHelper.GetWebResourceDetailMdPath(content.solution.UniqueName, wr.Name);
                string detailFullPath = Path.Combine(content.folderPath, detailRelPath);

                MdDocument doc = new MdDocument();
                string displayName = wr.DisplayName ?? wr.Name ?? "";
                doc.Root.Add(new MdHeading(displayName, 1));

                List<MdTableRow> propRows = new List<MdTableRow>();
                propRows.Add(new MdTableRow("Internal Name", wr.Name ?? ""));
                propRows.Add(new MdTableRow("Type", wr.GetTypeDisplayName()));
                propRows.Add(new MdTableRow("Introduced Version", wr.IntroducedVersion ?? ""));
                propRows.Add(new MdTableRow("Is Customizable", wr.IsCustomizable ? "Yes" : "No"));
                propRows.Add(new MdTableRow("Is Hidden", wr.IsHidden ? "Yes" : "No"));
                if (wr.Content != null)
                    propRows.Add(new MdTableRow("File Size", FormatFileSize(wr.Content.Length)));
                doc.Root.Add(new MdTable(new MdTableRow("Property", "Value"), propRows));

                // Source code
                if (wr.Content != null)
                {
                    doc.Root.Add(new MdHeading("Source Code", 2));
                    string textContent = Encoding.UTF8.GetString(wr.Content);
                    if (textContent.Length > 50000)
                    {
                        textContent = textContent.Substring(0, 50000) + "\n... (truncated)";
                    }
                    doc.Root.Add(new MdCodeBlock(textContent, GetCodeLanguage(wr)));
                }

                // SVG preview
                if (wr.IsImageType())
                {
                    doc.Root.Add(new MdHeading("Preview", 2));
                    string safeFileName = CharsetHelper.GetSafeName(wr.Name) + wr.GetFileExtension();
                    doc.Root.Add(new MdParagraph(new MdImageSpan(displayName, safeFileName)));
                }

                // Back link
                doc.Root.Add(new MdParagraph(new MdLinkSpan("\u2190 Back to Web Resources", "../" + indexFileName)));

                doc.Save(detailFullPath);
            }
        }

        private static string GetCodeLanguage(WebResourceEntity wr)
        {
            return wr.WebResourceType switch
            {
                "1" => "html",
                "2" => "css",
                "3" => "javascript",
                "4" => "xml",
                "9" => "xml",
                "11" => "xml",
                "12" => "xml",
                _ => ""
            };
        }

        private static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("F1") + " KB";
            return (bytes / (1024.0 * 1024.0)).ToString("F1") + " MB";
        }
    }
}
