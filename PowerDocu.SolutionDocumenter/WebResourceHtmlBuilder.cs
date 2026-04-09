using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using PowerDocu.Common;

namespace PowerDocu.SolutionDocumenter
{
    class WebResourceHtmlBuilder : HtmlBuilder
    {
        private readonly SolutionDocumentationContent content;
        private readonly string indexFileName;

        public WebResourceHtmlBuilder(SolutionDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            WriteDefaultStylesheet(content.folderPath);

            indexFileName = CrossDocLinkHelper.GetWebResourceDocHtmlPath(content.solution.UniqueName);

            List<WebResourceEntity> webResources = content.solution.Customizations.getWebResources();
            if (webResources.Count == 0) return;

            // Extract web resource file content from the solution ZIP
            ExtractWebResourceContent(webResources);

            // Build the index page
            StringBuilder body = new StringBuilder();
            BuildIndexPage(body, webResources);

            SaveHtmlFile(Path.Combine(content.folderPath, indexFileName),
                WrapInHtmlPage("Web Resources - " + content.solution.UniqueName, body.ToString(), GetNavigationHtml(webResources)));

            // Build individual detail pages for text-based resources
            BuildDetailPages(webResources);

            NotificationHelper.SendNotification("Created HTML documentation for web resources in " + content.solution.UniqueName);
        }

        private void ExtractWebResourceContent(List<WebResourceEntity> webResources)
        {
            string zipPath = content.context?.SourceZipPath;
            if (string.IsNullOrEmpty(zipPath) || !File.Exists(zipPath)) return;

            // Create output folder for extracted files
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

                // Write image files to disk so they can be referenced from HTML
                if (wr.IsImageType())
                {
                    string safeFileName = CharsetHelper.GetSafeName(wr.Name) + wr.GetFileExtension();
                    string outputFile = Path.Combine(wrOutputDir, safeFileName);
                    File.WriteAllBytes(outputFile, wr.Content);
                }
            }
        }

        private void BuildIndexPage(StringBuilder body, List<WebResourceEntity> webResources)
        {
            body.AppendLine(Heading(1, "Web Resources"));
            body.AppendLine(Paragraph($"This solution contains {webResources.Count} web resource(s)."));

            // Overview table
            body.AppendLine(HeadingWithId(2, "Overview", "overview"));
            body.Append(TableStart("Display Name", "Name", "Type", "Version"));
            foreach (WebResourceEntity wr in webResources.OrderBy(w => w.DisplayName ?? w.Name))
            {
                string displayName = wr.DisplayName ?? wr.Name ?? "";
                if (wr.IsTextType() && wr.Content != null)
                {
                    string detailPath = CrossDocLinkHelper.GetWebResourceDetailHtmlPath(content.solution.UniqueName, wr.Name);
                    displayName = Link(displayName, detailPath);
                    body.Append(TableRowRaw(displayName, Encode(wr.Name ?? ""), Encode(wr.GetTypeDisplayName()), Encode(wr.IntroducedVersion ?? "")));
                }
                else
                {
                    body.Append(TableRow(displayName, wr.Name ?? "", wr.GetTypeDisplayName(), wr.IntroducedVersion ?? ""));
                }
            }
            body.AppendLine(TableEnd());

            // Images section
            var imageResources = webResources.Where(w => w.IsImageType() && w.Content != null).OrderBy(w => w.DisplayName ?? w.Name).ToList();
            if (imageResources.Count > 0)
            {
                body.AppendLine(HeadingWithId(2, "Images", "images"));
                foreach (WebResourceEntity wr in imageResources)
                {
                    string displayName = wr.DisplayName ?? wr.Name ?? "";
                    string anchorId = SanitizeAnchorId("wr-" + wr.Name);
                    body.AppendLine(HeadingWithId(3, displayName, anchorId));

                    body.Append(TableStart("Property", "Value"));
                    body.Append(TableRow("Name", wr.Name ?? ""));
                    body.Append(TableRow("Type", wr.GetTypeDisplayName()));
                    body.Append(TableRow("Version", wr.IntroducedVersion ?? ""));
                    body.AppendLine(TableEnd());

                    // Display the image
                    string safeFileName = CharsetHelper.GetSafeName(wr.Name) + wr.GetFileExtension();
                    string imgPath = "WebResources/" + safeFileName;
                    body.AppendLine($"<div class=\"wr-image\" style=\"margin: 12px 0; padding: 12px; border: 1px solid #ddd; border-radius: 4px; background: #f9f9f9;\">");
                    body.AppendLine($"  <img src=\"{Encode(imgPath)}\" alt=\"{Encode(displayName)}\" style=\"max-width: 100%; max-height: 400px;\" />");
                    body.AppendLine("</div>");
                }
            }

            // Text-based resources section (links to detail pages)
            var textResources = webResources.Where(w => w.IsTextType() && w.Content != null && !w.IsImageType()).OrderBy(w => w.DisplayName ?? w.Name).ToList();
            if (textResources.Count > 0)
            {
                body.AppendLine(HeadingWithId(2, "Scripts and Markup", "scripts-and-markup"));
                body.Append(TableStart("Name", "Type", "Size"));
                foreach (WebResourceEntity wr in textResources)
                {
                    string displayName = wr.DisplayName ?? wr.Name ?? "";
                    string detailPath = CrossDocLinkHelper.GetWebResourceDetailHtmlPath(content.solution.UniqueName, wr.Name);
                    string nameLink = Link(displayName, detailPath);
                    string size = wr.Content != null ? FormatFileSize(wr.Content.Length) : "";
                    body.Append(TableRowRaw(nameLink, Encode(wr.GetTypeDisplayName()), Encode(size)));
                }
                body.AppendLine(TableEnd());
            }
        }

        private void BuildDetailPages(List<WebResourceEntity> webResources)
        {
            string wrOutputDir = Path.Combine(content.folderPath, "WebResources");
            Directory.CreateDirectory(wrOutputDir);
            WriteDefaultStylesheet(wrOutputDir);

            foreach (WebResourceEntity wr in webResources.Where(w => w.IsTextType() && w.Content != null))
            {
                string detailRelPath = CrossDocLinkHelper.GetWebResourceDetailHtmlPath(content.solution.UniqueName, wr.Name);
                string detailFullPath = Path.Combine(content.folderPath, detailRelPath);

                StringBuilder body = new StringBuilder();
                string displayName = wr.DisplayName ?? wr.Name ?? "";
                body.AppendLine(Heading(1, displayName));

                body.Append(TableStart("Property", "Value"));
                body.Append(TableRow("Internal Name", wr.Name ?? ""));
                body.Append(TableRow("Type", wr.GetTypeDisplayName()));
                body.Append(TableRow("Introduced Version", wr.IntroducedVersion ?? ""));
                body.Append(TableRow("Is Customizable", wr.IsCustomizable ? "Yes" : "No"));
                body.Append(TableRow("Is Hidden", wr.IsHidden ? "Yes" : "No"));
                if (wr.Content != null)
                    body.Append(TableRow("File Size", FormatFileSize(wr.Content.Length)));
                body.AppendLine(TableEnd());

                // Show file content
                if (wr.Content != null)
                {
                    body.AppendLine(HeadingWithId(2, "Source Code", "source-code"));
                    string textContent = Encoding.UTF8.GetString(wr.Content);
                    body.AppendLine(PreCodeBlock(textContent));
                }

                // If it's also an image type (SVG), show the rendered image too
                if (wr.IsImageType())
                {
                    body.AppendLine(HeadingWithId(2, "Preview", "preview"));
                    string safeFileName = CharsetHelper.GetSafeName(wr.Name) + wr.GetFileExtension();
                    string imgPath = safeFileName;
                    body.AppendLine($"<div style=\"margin: 12px 0; padding: 12px; border: 1px solid #ddd; border-radius: 4px; background: #f9f9f9;\">");
                    body.AppendLine($"  <img src=\"{Encode(imgPath)}\" alt=\"{Encode(displayName)}\" style=\"max-width: 100%; max-height: 400px;\" />");
                    body.AppendLine("</div>");
                }

                // Navigation for detail page
                string navHtml = GetDetailNavigationHtml(wr);

                SaveHtmlFile(detailFullPath,
                    WrapInHtmlPage(displayName + " - Web Resource", body.ToString(), navHtml, "../style.css"));
            }
        }

        private string GetNavigationHtml(List<WebResourceEntity> webResources)
        {
            var navItems = new List<(string label, string href, int level)>
            {
                ("← Solution", CrossDocLinkHelper.GetSolutionDocHtmlPath(content.solution.UniqueName), 0),
                ("Overview", indexFileName + "#overview", 0)
            };

            var imageResources = webResources.Where(w => w.IsImageType() && w.Content != null).OrderBy(w => w.DisplayName ?? w.Name).ToList();
            if (imageResources.Count > 0)
            {
                navItems.Add(("Images", indexFileName + "#images", 0));
                foreach (WebResourceEntity wr in imageResources)
                {
                    navItems.Add((wr.DisplayName ?? wr.Name ?? "", indexFileName + "#" + SanitizeAnchorId("wr-" + wr.Name), 1));
                }
            }

            var textResources = webResources.Where(w => w.IsTextType() && w.Content != null && !w.IsImageType()).OrderBy(w => w.DisplayName ?? w.Name).ToList();
            if (textResources.Count > 0)
            {
                navItems.Add(("Scripts and Markup", indexFileName + "#scripts-and-markup", 0));
                foreach (WebResourceEntity wr in textResources)
                {
                    string detailPath = CrossDocLinkHelper.GetWebResourceDetailHtmlPath(content.solution.UniqueName, wr.Name);
                    navItems.Add((wr.DisplayName ?? wr.Name ?? "", detailPath, 1));
                }
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"<div class=\"nav-title\">{Encode("Web Resources")}</div>");
            sb.Append(NavigationList(navItems));
            return sb.ToString();
        }

        private string GetDetailNavigationHtml(WebResourceEntity wr)
        {
            var navItems = new List<(string label, string href, int level)>
            {
                ("← Web Resources", "../" + indexFileName, 0),
                ("← Solution", "../" + CrossDocLinkHelper.GetSolutionDocHtmlPath(content.solution.UniqueName), 0),
                ("Properties", "#", 0),
                ("Source Code", "#source-code", 0)
            };

            if (wr.IsImageType())
            {
                navItems.Add(("Preview", "#preview", 0));
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"<div class=\"nav-title\">{Encode(wr.DisplayName ?? wr.Name ?? "")}</div>");
            sb.Append(NavigationList(navItems));
            return sb.ToString();
        }

        private static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("F1") + " KB";
            return (bytes / (1024.0 * 1024.0)).ToString("F1") + " MB";
        }
    }
}
