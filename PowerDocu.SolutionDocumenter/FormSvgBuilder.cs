using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using PowerDocu.Common;
using Svg;

namespace PowerDocu.SolutionDocumenter
{
    /// <summary>
    /// Generates SVG wireframe mockups for Dataverse forms, showing tabs, columns, sections, and controls.
    /// </summary>
    public static class FormSvgBuilder
    {
        // Layout constants
        private const int FormWidth = 800;
        private const int TabBarHeight = 36;
        private const int TabPaddingH = 16;
        private const int TabSpacing = 4;
        private const int TabFontSize = 12;
        private const int ColumnGap = 8;
        private const int SectionPadding = 8;
        private const int SectionHeaderHeight = 22;
        private const int SectionSpacing = 8;
        private const int ControlRowHeight = 28;
        private const int ControlPaddingV = 4;
        private const int FormPadding = 12;
        private const int FontSize = 11;
        private const int SmallFontSize = 9;

        // Colors matching Dataverse form styling
        private const string FormBorderColor = "#d2d0ce";
        private const string FormBgColor = "#ffffff";
        private const string TabBarBgColor = "#f3f2f1";
        private const string TabActiveBgColor = "#ffffff";
        private const string TabActiveTextColor = "#0078d4";
        private const string TabInactiveTextColor = "#605e5c";
        private const string TabActiveBorderColor = "#0078d4";
        private const string SectionBorderColor = "#edebe9";
        private const string SectionHeaderBgColor = "#faf9f8";
        private const string SectionHeaderTextColor = "#323130";
        private const string ControlBgColor = "#ffffff";
        private const string ControlBorderColor = "#8a8886";
        private const string LabelTextColor = "#605e5c";
        private const string ControlTextColor = "#323130";
        private const string HiddenOverlayColor = "#f3f2f1";
        private const string HiddenTextColor = "#a19f9d";
        private const string FormHeaderColor = "#0078d4";
        private const string DisabledBgColor = "#f3f2f1";
        private const string DisabledBorderColor = "#c8c6c4";
        private const string HiddenControlBgColor = "#faf9f8";

        // Cache to avoid regenerating the same SVG when multiple output formats are produced
        private static readonly Dictionary<FormEntity, FormSvgResult> _svgCache = new Dictionary<FormEntity, FormSvgResult>();
        private static readonly HashSet<string> _writtenFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Holds pre-computed SVG content and dimensions for a form.
        /// </summary>
        public class FormSvgResult
        {
            public string SvgContent { get; }
            public int Width { get; }
            public int Height { get; }

            public FormSvgResult(string svgContent, int width, int height)
            {
                SvgContent = svgContent;
                Width = width;
                Height = height;
            }
        }

        /// <summary>
        /// Clears the internal SVG cache. Call after all output formats have been generated to free memory.
        /// </summary>
        public static void ClearCache()
        {
            _svgCache.Clear();
            _writtenFiles.Clear();
        }

        /// <summary>
        /// Returns the cached SVG result for a form, or builds and caches it on first access.
        /// </summary>
        private static FormSvgResult GetOrBuild(FormEntity form, string tableDisplayName, Dictionary<string, string> columnDisplayNames)
        {
            if (_svgCache.TryGetValue(form, out FormSvgResult cached))
                return cached;

            var result = BuildFormSvg(form, tableDisplayName, columnDisplayNames);
            _svgCache[form] = result;
            return result;
        }

        /// <summary>
        /// Generates SVG files for all forms in the given table entity.
        /// Returns a dictionary mapping form name to the SVG filename (relative to folderPath).
        /// </summary>
        /// <param name="columnDisplayNames">Dictionary mapping logical column names to display names.</param>
        public static Dictionary<string, string> GenerateFormSvgs(TableEntity tableEntity, string folderPath, Dictionary<string, string> columnDisplayNames = null)
        {
            var result = new Dictionary<string, string>();
            string dataversePath = Path.Combine(folderPath, "Dataverse");
            Directory.CreateDirectory(dataversePath);

            foreach (FormEntity form in tableEntity.GetForms())
            {
                List<FormTab> tabs = form.GetTabs();
                if (tabs.Count == 0) continue;

                string formTypeLabel = form.GetFormTypeDisplayName();
                string safeName = CharsetHelper.GetSafeName(
                    tableEntity.getName() + "-form-" + formTypeLabel + "-" + form.GetFormName())
                    .Replace(" ", "-");
                string filename = safeName + ".svg";
                FormSvgResult svgResult = GetOrBuild(form, tableEntity.getLocalizedName(), columnDisplayNames);
                string fullPath = Path.Combine(dataversePath, filename);
                if (_writtenFiles.Add(fullPath))
                {
                    File.WriteAllText(fullPath, svgResult.SvgContent, Encoding.UTF8);
                    //generating the PNG from the SVG
                    var svgDocument = SvgDocument.Open(fullPath);
                    using (var bitmap = svgDocument.Draw())
                    {
                        bitmap?.Save(Path.ChangeExtension(fullPath, ".png"));
                    }
                }
                string formKey = form.GetFormName() + "|" + formTypeLabel;
                result[formKey] = "Dataverse/" + filename;
            }

            return result;
        }

        /// <summary>
        /// Generates SVG content for a single form as a string (for inline embedding).
        /// </summary>
        /// <param name="columnDisplayNames">Dictionary mapping logical column names to display names.</param>
        public static string GenerateFormSvgContent(FormEntity form, string tableDisplayName, Dictionary<string, string> columnDisplayNames = null)
        {
            List<FormTab> tabs = form.GetTabs();
            if (tabs.Count == 0) return "";
            return GetOrBuild(form, tableDisplayName, columnDisplayNames).SvgContent;
        }

        private static FormSvgResult BuildFormSvg(FormEntity form, string tableDisplayName, Dictionary<string, string> columnDisplayNames)
        {
            List<FormTab> tabs = form.GetTabs();
            var sb = new StringBuilder();

            // First pass: measure the total height needed
            // We render all tabs stacked vertically so the full form layout is visible
            int y = 0;

            // Form header
            int formHeaderHeight = 30;
            y += formHeaderHeight;

            // Tab sections (one block per tab, each with tab-bar + content)
            var tabMeasurements = new List<(FormTab tab, int contentHeight, List<TabColumnMeasurement> columns)>();
            foreach (FormTab tab in tabs)
            {
                var measurement = MeasureTab(tab);
                tabMeasurements.Add((tab, measurement.totalHeight, measurement.columns));
                y += TabBarHeight + measurement.totalHeight + FormPadding;
            }

            int totalHeight = y + FormPadding;
            int totalWidth = FormWidth + 2;

            // Start SVG
            sb.AppendLine($"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{totalWidth}\" height=\"{totalHeight}\" viewBox=\"0 0 {totalWidth} {totalHeight}\" font-family=\"Segoe UI, Helvetica, Arial, sans-serif\">");
            sb.AppendLine("<defs>");
            sb.AppendLine("  <style>");
            sb.AppendLine("    .form-label { font-size: " + SmallFontSize + "px; fill: " + LabelTextColor + "; }");
            sb.AppendLine("    .form-control-text { font-size: " + FontSize + "px; fill: " + ControlTextColor + "; }");
            sb.AppendLine("    .section-header-text { font-size: " + FontSize + "px; fill: " + SectionHeaderTextColor + "; font-weight: 600; }");
            sb.AppendLine("    .tab-text-active { font-size: " + TabFontSize + "px; fill: " + TabActiveTextColor + "; font-weight: 600; }");
            sb.AppendLine("    .tab-text-inactive { font-size: " + TabFontSize + "px; fill: " + TabInactiveTextColor + "; }");
            sb.AppendLine("    .form-header-text { font-size: 14px; fill: #ffffff; font-weight: 600; }");
            sb.AppendLine("    .hidden-text { font-size: " + SmallFontSize + "px; fill: " + HiddenTextColor + "; font-style: italic; }");
            sb.AppendLine("  </style>");
            sb.AppendLine("</defs>");

            int currentY = 0;

            // Form header bar
            sb.AppendLine($"<rect x=\"1\" y=\"{currentY}\" width=\"{FormWidth}\" height=\"{formHeaderHeight}\" rx=\"4\" ry=\"4\" fill=\"{FormHeaderColor}\" />");
            sb.AppendLine($"<text x=\"{FormPadding + 8}\" y=\"{currentY + 20}\" class=\"form-header-text\">{Escape(form.GetFormName())} — {Escape(tableDisplayName)} ({Escape(form.GetFormTypeDisplayName())})</text>");
            currentY += formHeaderHeight;

            // Render each tab
            for (int t = 0; t < tabs.Count; t++)
            {
                var (tab, contentHeight, columns) = tabMeasurements[t];
                RenderTabBlock(sb, tabs, t, currentY, contentHeight, columns, columnDisplayNames);
                currentY += TabBarHeight + contentHeight + FormPadding;
            }

            sb.AppendLine("</svg>");
            return new FormSvgResult(sb.ToString(), totalWidth, totalHeight);
        }

        private static void RenderTabBlock(StringBuilder sb, List<FormTab> allTabs, int activeTabIndex,
            int startY, int contentHeight, List<TabColumnMeasurement> columns, Dictionary<string, string> columnDisplayNames)
        {
            FormTab activeTab = allTabs[activeTabIndex];
            int totalBlockHeight = TabBarHeight + contentHeight;

            // Tab bar background
            sb.AppendLine($"<rect x=\"1\" y=\"{startY}\" width=\"{FormWidth}\" height=\"{TabBarHeight}\" fill=\"{TabBarBgColor}\" stroke=\"{FormBorderColor}\" stroke-width=\"1\" />");

            // Tab labels
            int tabX = TabSpacing + 1;
            for (int i = 0; i < allTabs.Count; i++)
            {
                string tabName = allTabs[i].GetName();
                if (!allTabs[i].IsVisible()) tabName += " 👁";
                int estimatedWidth = tabName.Length * 7 + TabPaddingH * 2;
                bool isActive = (i == activeTabIndex);

                if (isActive)
                {
                    // Active tab with bottom border highlight
                    sb.AppendLine($"<rect x=\"{tabX}\" y=\"{startY}\" width=\"{estimatedWidth}\" height=\"{TabBarHeight}\" fill=\"{TabActiveBgColor}\" />");
                    sb.AppendLine($"<rect x=\"{tabX}\" y=\"{startY + TabBarHeight - 3}\" width=\"{estimatedWidth}\" height=\"3\" fill=\"{TabActiveBorderColor}\" />");
                    sb.AppendLine($"<text x=\"{tabX + TabPaddingH}\" y=\"{startY + TabBarHeight / 2 + 4}\" class=\"tab-text-active\">{Escape(tabName)}</text>");
                }
                else
                {
                    sb.AppendLine($"<text x=\"{tabX + TabPaddingH}\" y=\"{startY + TabBarHeight / 2 + 4}\" class=\"tab-text-inactive\">{Escape(tabName)}</text>");
                }
                tabX += estimatedWidth + TabSpacing;
            }

            // Content area
            int contentY = startY + TabBarHeight;
            sb.AppendLine($"<rect x=\"1\" y=\"{contentY}\" width=\"{FormWidth}\" height=\"{contentHeight}\" fill=\"{FormBgColor}\" stroke=\"{FormBorderColor}\" stroke-width=\"1\" />");

            // Hidden overlay for the entire tab if not visible
            if (!activeTab.IsVisible())
            {
                sb.AppendLine($"<rect x=\"1\" y=\"{contentY}\" width=\"{FormWidth}\" height=\"{contentHeight}\" fill=\"{HiddenOverlayColor}\" fill-opacity=\"0.5\" />");
                sb.AppendLine($"<text x=\"{FormWidth / 2}\" y=\"{contentY + 16}\" text-anchor=\"middle\" class=\"hidden-text\">(hidden tab)</text>");
            }

            // Render columns side-by-side
            if (columns.Count == 0) return;

            int availableWidth = FormWidth - FormPadding * 2 - (columns.Count - 1) * ColumnGap;
            int columnWidth = availableWidth / columns.Count;

            for (int c = 0; c < columns.Count; c++)
            {
                int colX = 1 + FormPadding + c * (columnWidth + ColumnGap);
                int colY = contentY + SectionPadding;

                foreach (var sectionMeasure in columns[c].sections)
                {
                    RenderSection(sb, sectionMeasure.section, colX, colY, columnWidth, sectionMeasure.height, columnDisplayNames);
                    colY += sectionMeasure.height + SectionSpacing;
                }
            }
        }

        private static void RenderSection(StringBuilder sb, FormSection section, int x, int y, int width, int height, Dictionary<string, string> columnDisplayNames)
        {
            // Section container
            sb.AppendLine($"<rect x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{height}\" rx=\"2\" ry=\"2\" fill=\"{FormBgColor}\" stroke=\"{SectionBorderColor}\" stroke-width=\"1\" />");

            // Section header
            sb.AppendLine($"<rect x=\"{x}\" y=\"{y}\" width=\"{width}\" height=\"{SectionHeaderHeight}\" rx=\"2\" ry=\"2\" fill=\"{SectionHeaderBgColor}\" />");
            string sectionLabel = section.GetName();
            if (!section.IsVisible()) sectionLabel += " (hidden)";
            sb.AppendLine($"<text x=\"{x + SectionPadding}\" y=\"{y + 15}\" class=\"section-header-text\">{Escape(sectionLabel)}</text>");

            if (!section.IsVisible())
            {
                sb.AppendLine($"<rect x=\"{x}\" y=\"{y + SectionHeaderHeight}\" width=\"{width}\" height=\"{height - SectionHeaderHeight}\" fill=\"{HiddenOverlayColor}\" fill-opacity=\"0.4\" />");
            }

            // Controls
            List<FormControl> controls = section.GetControls();
            int controlY = y + SectionHeaderHeight + ControlPaddingV;
            int controlAreaWidth = width - SectionPadding * 2;

            foreach (FormControl control in controls)
            {
                RenderControl(sb, control, x + SectionPadding, controlY, controlAreaWidth, columnDisplayNames);
                controlY += ControlRowHeight;
            }
        }

        private static void RenderControl(StringBuilder sb, FormControl control, int x, int y, int width, Dictionary<string, string> columnDisplayNames)
        {
            string label = ResolveControlLabel(control, columnDisplayNames);
            bool hasLabel = !control.IsLabelHidden();
            bool isDisabled = control.IsDisabled();
            bool isHidden = !control.IsVisible();

            // Choose colors based on state
            string bgColor = isDisabled ? DisabledBgColor : ControlBgColor;
            string borderColor = isDisabled ? DisabledBorderColor : ControlBorderColor;

            if (isHidden)
            {
                bgColor = HiddenControlBgColor;
                borderColor = DisabledBorderColor;
            }

            // Suffix for hidden controls
            string labelSuffix = isHidden ? " (hidden)" : "";

            if (hasLabel)
            {
                // Label above control input
                string labelClass = isHidden ? "hidden-text" : "form-label";
                sb.AppendLine($"<text x=\"{x + 4}\" y=\"{y + 10}\" class=\"{labelClass}\">{Escape(label + labelSuffix)}</text>");
                // Control input box
                int inputY = y + 12;
                int inputHeight = 14;
                sb.AppendLine($"<rect x=\"{x + 2}\" y=\"{inputY}\" width=\"{width - 4}\" height=\"{inputHeight}\" rx=\"2\" ry=\"2\" fill=\"{bgColor}\" stroke=\"{borderColor}\" stroke-width=\"0.5\" />");
                if (isDisabled && !isHidden)
                {
                    // Show lock icon indicator for disabled fields
                    sb.AppendLine($"<text x=\"{x + width - 16}\" y=\"{inputY + 11}\" style=\"font-size: 9px; fill: {HiddenTextColor};\">\uD83D\uDD12</text>");
                }
            }
            else
            {
                // Full-height control without label
                sb.AppendLine($"<rect x=\"{x + 2}\" y=\"{y + 2}\" width=\"{width - 4}\" height=\"{ControlRowHeight - 6}\" rx=\"2\" ry=\"2\" fill=\"{bgColor}\" stroke=\"{borderColor}\" stroke-width=\"0.5\" />");
                string textClass = isHidden ? "hidden-text" : "form-control-text";
                sb.AppendLine($"<text x=\"{x + 8}\" y=\"{y + 16}\" class=\"{textClass}\">{Escape(label + labelSuffix)}</text>");
            }
        }

        #region Measurement

        private class TabMeasurement
        {
            public int totalHeight;
            public List<TabColumnMeasurement> columns;
        }

        private class TabColumnMeasurement
        {
            public List<SectionMeasurement> sections = new List<SectionMeasurement>();
            public int totalHeight;
        }

        private class SectionMeasurement
        {
            public FormSection section;
            public int height;
        }

        private static TabMeasurement MeasureTab(FormTab tab)
        {
            var result = new TabMeasurement();
            result.columns = new List<TabColumnMeasurement>();

            List<FormTabColumn> formColumns = tab.GetColumns();
            if (formColumns.Count == 0)
            {
                // Fallback: treat all sections as a single column
                var colMeasure = new TabColumnMeasurement();
                foreach (FormSection section in tab.GetSections())
                {
                    int sectionHeight = MeasureSection(section);
                    colMeasure.sections.Add(new SectionMeasurement { section = section, height = sectionHeight });
                    colMeasure.totalHeight += sectionHeight + SectionSpacing;
                }
                result.columns.Add(colMeasure);
            }
            else
            {
                foreach (FormTabColumn formColumn in formColumns)
                {
                    var colMeasure = new TabColumnMeasurement();
                    foreach (FormSection section in formColumn.GetSections())
                    {
                        int sectionHeight = MeasureSection(section);
                        colMeasure.sections.Add(new SectionMeasurement { section = section, height = sectionHeight });
                        colMeasure.totalHeight += sectionHeight + SectionSpacing;
                    }
                    result.columns.Add(colMeasure);
                }
            }

            // The tab height is the max column height, plus padding
            result.totalHeight = result.columns.Count > 0
                ? result.columns.Max(c => c.totalHeight) + SectionPadding * 2
                : SectionPadding * 2;

            // Ensure minimum height
            if (result.totalHeight < 40) result.totalHeight = 40;

            return result;
        }

        private static int MeasureSection(FormSection section)
        {
            int controlCount = section.GetControls().Count;
            int controlsHeight = controlCount * ControlRowHeight + ControlPaddingV * 2;
            return SectionHeaderHeight + Math.Max(controlsHeight, ControlPaddingV * 2);
        }

        #endregion

        /// <summary>
        /// Returns the dimensions (width, height) of the SVG that would be generated for the given form.
        /// Useful for Word document embedding.
        /// </summary>
        public static (int width, int height) MeasureFormSvg(FormEntity form)
        {
            // Use cached dimensions if available (avoids redundant measurement pass)
            if (_svgCache.TryGetValue(form, out FormSvgResult cached))
                return (cached.Width, cached.Height);

            List<FormTab> tabs = form.GetTabs();
            if (tabs.Count == 0) return (0, 0);

            int y = 30; // form header
            foreach (FormTab tab in tabs)
            {
                var measurement = MeasureTab(tab);
                y += TabBarHeight + measurement.totalHeight + FormPadding;
            }

            return (FormWidth + 2, y + FormPadding);
        }

        /// <summary>
        /// Resolves the display label for a control, using the column display name dictionary
        /// to show "Display Name (internal_name)" format when available.
        /// </summary>
        private static string ResolveControlLabel(FormControl control, Dictionary<string, string> columnDisplayNames)
        {
            string fieldName = control.GetDataFieldName();
            if (!string.IsNullOrEmpty(fieldName) && columnDisplayNames != null)
            {
                if (columnDisplayNames.TryGetValue(fieldName, out string displayName) && !string.IsNullOrEmpty(displayName))
                {
                    return displayName + " (" + fieldName + ")";
                }
            }
            return control.GetDisplayLabel();
        }

        private static string Escape(string text)
        {
            return HttpUtility.HtmlEncode(text ?? "");
        }
    }
}
