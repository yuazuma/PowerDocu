using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using PowerDocu.Common;

namespace PowerDocu.AppModuleDocumenter
{
    /// <summary>
    /// Generates an SVG visualization of a Model-Driven App site map navigation structure.
    /// Renders Areas as top-level columns, Groups as card headers, and SubAreas as two-line list items
    /// (title on the first line, target detail on the second line).
    /// </summary>
    public static class SiteMapSvgBuilder
    {
        // Layout constants
        private const int CanvasPadding = 20;
        private const int AreaSpacing = 16;
        private const int AreaMinWidth = 240;
        private const int GroupSpacing = 12;
        private const int SubAreaRowHeight = 38; // two lines: title + detail
        private const int GroupHeaderHeight = 28;
        private const int GroupPadding = 8;
        private const int AreaHeaderHeight = 36;
        private const int AreaPadding = 10;
        private const int FontSize = 12;
        private const int SmallFontSize = 11;
        private const int DetailFontSize = 9;
        private const int TitleFontSize = 14;

        // Approximate pixels per character for width estimation
        private const double CharsPerPixelTitle = 6.5;
        private const double CharsPerPixelDetail = 5.5;
        private const int IconAndPaddingWidth = 28; // arrow icon + left padding

        // Colors
        private const string BgColor = "#ffffff";
        private const string AreaHeaderBg = "#0078d4";
        private const string AreaHeaderText = "#ffffff";
        private const string AreaBorderColor = "#d2d0ce";
        private const string GroupHeaderBg = "#f3f2f1";
        private const string GroupHeaderText = "#323130";
        private const string GroupBorderColor = "#edebe9";
        private const string SubAreaText = "#323130";
        private const string SubAreaDetailText = "#605e5c";

        /// <summary>
        /// Generates a site map SVG file for the given site map entity.
        /// </summary>
        public static void GenerateSiteMapSvg(AppModuleSiteMap siteMap, string folderPath, string filename)
        {
            if (siteMap == null || siteMap.Areas.Count == 0) return;

            Directory.CreateDirectory(folderPath);
            string svgContent = BuildSvg(siteMap);
            string safeFilename = "sitemap-" + filename.Replace(" ", "-") + ".svg";
            File.WriteAllText(Path.Combine(folderPath, safeFilename), svgContent, Encoding.UTF8);
        }

        /// <summary>
        /// Estimates the pixel width needed for the widest text in an area column.
        /// </summary>
        private static int EstimateAreaWidth(SiteMapArea area)
        {
            int maxWidth = 0;
            foreach (var group in area.Groups)
            {
                // Group header
                int groupTitleWidth = (int)(group.Title.Length * CharsPerPixelTitle) + GroupPadding * 2;
                maxWidth = Math.Max(maxWidth, groupTitleWidth);

                foreach (var subArea in group.SubAreas)
                {
                    string title = !string.IsNullOrEmpty(subArea.Title) ? subArea.Title : subArea.Id;
                    string detail = subArea.GetTargetDescription();
                    int titleWidth = (int)(title.Length * CharsPerPixelTitle) + IconAndPaddingWidth;
                    int detailWidth = !string.IsNullOrEmpty(detail) ? (int)(detail.Length * CharsPerPixelDetail) + IconAndPaddingWidth : 0;
                    maxWidth = Math.Max(maxWidth, Math.Max(titleWidth, detailWidth));
                }
            }
            // Add area padding on both sides
            return Math.Max(AreaMinWidth, maxWidth + AreaPadding * 2 + GroupPadding * 2);
        }

        private static string BuildSvg(AppModuleSiteMap siteMap)
        {
            var sb = new StringBuilder();

            // Measure each area column
            var areaMeasurements = siteMap.Areas.Select(area =>
            {
                int height = AreaHeaderHeight + AreaPadding;
                foreach (var group in area.Groups)
                {
                    height += GroupHeaderHeight;
                    height += group.SubAreas.Count * SubAreaRowHeight;
                    height += GroupPadding * 2 + GroupSpacing;
                }
                int width = EstimateAreaWidth(area);
                return new { Area = area, Height = height, Width = width };
            }).ToList();

            int totalHeight = areaMeasurements.Max(a => a.Height) + CanvasPadding * 2;
            int totalWidth = CanvasPadding * 2 + areaMeasurements.Sum(a => a.Width) + (areaMeasurements.Count - 1) * AreaSpacing;

            // SVG header
            sb.AppendLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 {totalWidth} {totalHeight}\" width=\"{totalWidth}\" height=\"{totalHeight}\">");
            sb.AppendLine("<defs>");
            sb.AppendLine("  <style>");
            sb.AppendLine($"    .area-header {{ font-family: 'Segoe UI', sans-serif; font-size: {TitleFontSize}px; font-weight: 600; fill: {AreaHeaderText}; }}");
            sb.AppendLine($"    .group-header {{ font-family: 'Segoe UI', sans-serif; font-size: {FontSize}px; font-weight: 600; fill: {GroupHeaderText}; }}");
            sb.AppendLine($"    .subarea-title {{ font-family: 'Segoe UI', sans-serif; font-size: {SmallFontSize}px; fill: {SubAreaText}; }}");
            sb.AppendLine($"    .subarea-detail {{ font-family: 'Segoe UI', sans-serif; font-size: {DetailFontSize}px; fill: {SubAreaDetailText}; }}");
            sb.AppendLine("  </style>");
            sb.AppendLine("</defs>");

            // Background
            sb.AppendLine($"<rect width=\"{totalWidth}\" height=\"{totalHeight}\" fill=\"{BgColor}\" rx=\"4\" />");

            int x = CanvasPadding;
            foreach (var am in areaMeasurements)
            {
                var area = am.Area;
                int areaWidth = am.Width;
                int y = CanvasPadding;

                // Area border
                sb.AppendLine($"<rect x=\"{x}\" y=\"{y}\" width=\"{areaWidth}\" height=\"{am.Height}\" rx=\"4\" fill=\"none\" stroke=\"{AreaBorderColor}\" stroke-width=\"1\" />");

                // Area header
                sb.AppendLine($"<rect x=\"{x}\" y=\"{y}\" width=\"{areaWidth}\" height=\"{AreaHeaderHeight}\" rx=\"4\" fill=\"{AreaHeaderBg}\" />");
                sb.AppendLine($"<rect x=\"{x}\" y=\"{y + AreaHeaderHeight - 4}\" width=\"{areaWidth}\" height=\"4\" fill=\"{AreaHeaderBg}\" />");
                sb.AppendLine($"<text x=\"{x + AreaPadding}\" y=\"{y + AreaHeaderHeight / 2 + 5}\" class=\"area-header\">{Encode(area.Title)}</text>");
                y += AreaHeaderHeight + AreaPadding;

                foreach (var group in area.Groups)
                {
                    int groupX = x + AreaPadding;
                    int groupWidth = areaWidth - AreaPadding * 2;
                    int groupContentHeight = GroupHeaderHeight + group.SubAreas.Count * SubAreaRowHeight + GroupPadding;

                    // Group border
                    sb.AppendLine($"<rect x=\"{groupX}\" y=\"{y}\" width=\"{groupWidth}\" height=\"{groupContentHeight}\" rx=\"2\" fill=\"none\" stroke=\"{GroupBorderColor}\" stroke-width=\"1\" />");

                    // Group header background
                    sb.AppendLine($"<rect x=\"{groupX}\" y=\"{y}\" width=\"{groupWidth}\" height=\"{GroupHeaderHeight}\" rx=\"2\" fill=\"{GroupHeaderBg}\" />");
                    sb.AppendLine($"<rect x=\"{groupX}\" y=\"{y + GroupHeaderHeight - 2}\" width=\"{groupWidth}\" height=\"2\" fill=\"{GroupHeaderBg}\" />");
                    sb.AppendLine($"<text x=\"{groupX + GroupPadding}\" y=\"{y + GroupHeaderHeight / 2 + 4}\" class=\"group-header\">{Encode(group.Title)}</text>");
                    y += GroupHeaderHeight;

                    foreach (var subArea in group.SubAreas)
                    {
                        string title = !string.IsNullOrEmpty(subArea.Title) ? subArea.Title : subArea.Id;
                        string detail = subArea.GetTargetDescription();

                        // Row background
                        sb.AppendLine($"<rect x=\"{groupX + 1}\" y=\"{y}\" width=\"{groupWidth - 2}\" height=\"{SubAreaRowHeight}\" fill=\"{BgColor}\" />");

                        // Navigation icon (simple arrow)
                        int iconX = groupX + GroupPadding;
                        int titleY = y + 14; // first line baseline
                        sb.AppendLine($"<path d=\"M{iconX},{titleY - 4} L{iconX + 6},{titleY} L{iconX},{titleY + 4}\" fill=\"{SubAreaDetailText}\" />");

                        // Title text — first line
                        int textX = iconX + 12;
                        int maxTextWidth = groupWidth - GroupPadding * 2 - 12; // available width for text
                        string truncatedTitle = TruncateToFit(title, maxTextWidth, CharsPerPixelTitle);
                        sb.AppendLine($"<text x=\"{textX}\" y=\"{titleY}\" class=\"subarea-title\">{Encode(truncatedTitle)}</text>");

                        // Detail text — second line
                        if (!string.IsNullOrEmpty(detail))
                        {
                            int detailY = y + 28; // second line baseline
                            string truncatedDetail = TruncateToFit(detail, maxTextWidth, CharsPerPixelDetail);
                            sb.AppendLine($"<text x=\"{textX}\" y=\"{detailY}\" class=\"subarea-detail\">{Encode(truncatedDetail)}</text>");
                        }

                        y += SubAreaRowHeight;
                    }

                    y += GroupPadding + GroupSpacing;
                }

                x += areaWidth + AreaSpacing;
            }

            sb.AppendLine("</svg>");
            return sb.ToString();
        }

        /// <summary>
        /// Truncates text to fit within the given pixel width, appending "..." if needed.
        /// </summary>
        private static string TruncateToFit(string text, int maxPixelWidth, double charsPerPixel)
        {
            int maxChars = (int)(maxPixelWidth / charsPerPixel);
            if (maxChars < 4) maxChars = 4;
            if (text.Length <= maxChars) return text;
            return text.Substring(0, maxChars - 3) + "...";
        }

        private static string Encode(string text)
        {
            return HttpUtility.HtmlEncode(text ?? "");
        }
    }
}
