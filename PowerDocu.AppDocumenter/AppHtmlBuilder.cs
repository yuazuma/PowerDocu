using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using PowerDocu.Common;
using Svg;

namespace PowerDocu.AppDocumenter
{
    class AppHtmlBuilder : HtmlBuilder
    {
        private readonly AppDocumentationContent content;
        private readonly string mainFileName, appDetailsFileName, variablesFileName, dataSourcesFileName, resourcesFileName, controlsFileName;
        private readonly Dictionary<string, string> screenFileNames = new Dictionary<string, string>();
        private readonly Dictionary<string, string> datasourceFileNames = new Dictionary<string, string>();
        private readonly int appLogoWidth = 250;
        private bool documentChangedDefaultsOnly;
        private bool showDefaults;
        private bool documentSampleData;

        public AppHtmlBuilder(AppDocumentationContent contentdocumentation, bool documentChangedDefaultsOnly = false, bool showDefaults = true, bool documentSampleData = false)
        {
            content = contentdocumentation;
            this.documentChangedDefaultsOnly = documentChangedDefaultsOnly;
            this.showDefaults = showDefaults;
            this.documentSampleData = documentSampleData;
            Directory.CreateDirectory(content.folderPath);
            Directory.CreateDirectory(Path.Combine(content.folderPath, "resources"));
            WriteDefaultStylesheet(content.folderPath);

            mainFileName = ("index-" + content.filename + ".html").Replace(" ", "-");
            appDetailsFileName = ("appdetails-" + content.filename + ".html").Replace(" ", "-");
            variablesFileName = ("variables-" + content.filename + ".html").Replace(" ", "-");
            dataSourcesFileName = ("datasources-" + content.filename + ".html").Replace(" ", "-");
            resourcesFileName = ("resources-" + content.filename + ".html").Replace(" ", "-");
            controlsFileName = ("controls-" + content.filename + ".html").Replace(" ", "-");

            foreach (DataSource datasource in content.appDataSources.dataSources.OrderBy(o => o.Name).ToList())
            {
                datasourceFileNames[datasource.Name] = ("datasource-" + CharsetHelper.GetSafeName(datasource.Name) + "-" + content.filename + ".html").Replace(" ", "-");
            }
            foreach (ControlEntity screen in content.appControls.controls.Where(o => o.Type == "screen").OrderBy(o => o.Name).ToList())
            {
                screenFileNames[screen.Name] = ("screen-" + CharsetHelper.GetSafeName(screen.Name) + "-" + content.filename + ".html").Replace(" ", "-");
            }

            addAppMetadataAndOverview();
            addAppDetails();
            addAppVariablesInfo();
            addAppDataSources();
            addAppResources();
            addAppControlsOverview();
            addDetailedAppControls();
            NotificationHelper.SendNotification("Created HTML documentation for " + content.Name);
        }

        private string getNavigationHtml()
        {
            var navItems = new List<(string label, string href)>();
            if (content.context?.Solution != null)
            {
                if (content.context?.Config?.documentSolution == true)
                    navItems.Add(("Solution", "../" + CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName)));
                else
                    navItems.Add((content.context.Solution.UniqueName, ""));
            }
            navItems.AddRange(new (string label, string href)[]
            {
                ("Overview", mainFileName),
                ("App Details", appDetailsFileName),
                ("Variables", variablesFileName),
                ("DataSources", dataSourcesFileName),
                ("Resources", resourcesFileName),
                ("Controls", controlsFileName)
            });
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"<div class=\"nav-title\">{Encode(content.Name)}</div>");
            sb.Append(NavigationList(navItems));
            return sb.ToString();
        }

        private string buildMetadataTable()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(TableStart("Property", "Value"));
            sb.Append(TableRow("App Name", content.Name));
            if (!String.IsNullOrEmpty(content.appProperties.appLogo))
            {
                if (content.ResourceStreams.TryGetValue(content.appProperties.appLogo, out MemoryStream resourceStream))
                {
                    Bitmap appLogo;
                    if (!String.IsNullOrEmpty(content.appProperties.appBackgroundColour))
                    {
                        Color c = ColorTranslator.FromHtml(ColourHelper.ParseColor(content.appProperties.appBackgroundColour));
                        Bitmap bmp = new Bitmap(resourceStream);
                        appLogo = new Bitmap(bmp.Width, bmp.Height);
                        Rectangle rect = new Rectangle(Point.Empty, bmp.Size);
                        using (Graphics G = Graphics.FromImage(appLogo))
                        {
                            G.Clear(c);
                            G.DrawImageUnscaledAndClipped(bmp, rect);
                        }
                        appLogo.Save(content.folderPath + @"resources\applogo.png");
                    }
                    else
                    {
                        using Stream streamToWriteTo = File.Open(content.folderPath + @"resources\applogo.png", FileMode.Create);
                        resourceStream.CopyTo(streamToWriteTo);
                        resourceStream.Position = 0;
                        appLogo = new Bitmap(resourceStream);
                    }
                    resourceStream.Position = 0;
                    if (appLogo.Width > appLogoWidth)
                    {
                        Bitmap resized = new Bitmap(appLogo, new Size(appLogoWidth, appLogoWidth * appLogo.Height / appLogo.Width));
                        resized.Save(content.folderPath + @"resources\applogoSmall.png");
                        sb.Append(TableRowRaw("App Logo", Image("App Logo", "resources/applogoSmall.png")));
                    }
                    else
                    {
                        sb.Append(TableRowRaw("App Logo", Image("App Logo", "resources/applogo.png")));
                    }
                }
            }
            sb.Append(TableRow(content.appProperties.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion()));
            sb.Append(TableEnd());
            return sb.ToString();
        }

        private void addAppMetadataAndOverview()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.appProperties.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.appProperties.headerAppStatistics));
            body.Append(TableStart("Component Type", "Count"));
            foreach (KeyValuePair<string, string> kvp in content.appProperties.statisticsTable)
            {
                body.Append(TableRow(kvp.Key, kvp.Value));
            }
            body.AppendLine(TableEnd());
            SaveHtmlFile(Path.Combine(content.folderPath, mainFileName),
                WrapInHtmlPage(content.appProperties.header, body.ToString(), getNavigationHtml()));
        }

        private void addAppDetails()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.appProperties.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.appProperties.headerAppProperties));
            body.Append(TableStart("App Property", "Value"));
            foreach (Expression property in content.appProperties.appProperties)
            {
                if (!content.appProperties.propertiesToSkip.Contains(property.expressionOperator))
                {
                    body.Append(TableRow(property.expressionOperator, property.expressionOperands[0].ToString()));
                }
            }
            body.AppendLine(TableEnd());

            body.AppendLine(Heading(2, content.appProperties.headerAppInfo));
            body.Append(buildControlTable(content.appControls.controls.First<ControlEntity>(o => o.Type == "appinfo")));

            body.AppendLine(Heading(2, content.appProperties.headerAppPreviewFlags));
            if (content.appProperties.appPreviewsFlagProperty != null)
            {
                body.Append(TableStart("Preview Flag", "Value"));
                foreach (Expression flagProp in content.appProperties.appPreviewsFlagProperty.expressionOperands)
                {
                    body.Append(TableRow(flagProp.expressionOperator, flagProp.expressionOperands[0].ToString()));
                }
                body.AppendLine(TableEnd());
            }
            SaveHtmlFile(Path.Combine(content.folderPath, appDetailsFileName),
                WrapInHtmlPage("App Details - " + content.Name, body.ToString(), getNavigationHtml()));
        }

        private void addAppVariablesInfo()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.appProperties.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.appVariablesInfo.header));
            body.AppendLine(Paragraph(content.appVariablesInfo.infoText));

            // Global Variables
            body.AppendLine(Heading(3, content.appVariablesInfo.headerGlobalVariables));
            foreach (string var in content.appVariablesInfo.globalVariables)
            {
                body.AppendLine(Heading(4, var));
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out List<ControlPropertyReference> references);
                if (references != null)
                {
                    body.AppendLine(Paragraph("Variable used in:"));
                    body.Append(TableStart("Control", "Property"));
                    foreach (ControlPropertyReference reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        if (reference.Control.Type == "appinfo")
                        {
                            body.Append(TableRowRaw(
                                Link(reference.Control.Name, appDetailsFileName),
                                Encode(reference.RuleProperty)));
                        }
                        else
                        {
                            string screenFileName = screenFileNames.GetValueOrDefault(reference.Control.Screen()?.Name, "#");
                            string controlAnchor = SanitizeAnchorId(reference.Control.Name);
                            body.Append(TableRowRaw(
                                Link(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")", screenFileName + "#" + controlAnchor),
                                Encode(reference.RuleProperty)));
                        }
                    }
                    body.AppendLine(TableEnd());
                }
            }

            // Context Variables
            body.AppendLine(Heading(3, content.appVariablesInfo.headerContextVariables));
            foreach (string var in content.appVariablesInfo.contextVariables)
            {
                body.AppendLine(Heading(4, var));
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out List<ControlPropertyReference> references);
                if (references != null)
                {
                    body.AppendLine(Paragraph("Variable used in:"));
                    body.Append(TableStart("Control", "Property"));
                    foreach (ControlPropertyReference reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        if (reference.Control.Type == "appinfo")
                        {
                            body.Append(TableRowRaw(
                                Link(reference.Control.Name, appDetailsFileName),
                                Encode(reference.RuleProperty)));
                        }
                        else
                        {
                            string screenFileName = screenFileNames.GetValueOrDefault(reference.Control.Screen()?.Name, "#");
                            string controlAnchor = SanitizeAnchorId(reference.Control.Name);
                            body.Append(TableRowRaw(
                                Link(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")", screenFileName + "#" + controlAnchor),
                                Encode(reference.RuleProperty)));
                        }
                    }
                    body.AppendLine(TableEnd());
                }
            }

            // Collections
            body.AppendLine(Heading(3, content.appVariablesInfo.headerCollections));
            foreach (string var in content.appVariablesInfo.collections)
            {
                body.AppendLine(Heading(4, var));
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out List<ControlPropertyReference> references);
                if (references != null)
                {
                    body.AppendLine(Paragraph("Collection used in:"));
                    body.Append(TableStart("Control", "Property"));
                    foreach (ControlPropertyReference reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        if (reference.Control.Type == "appinfo")
                        {
                            body.Append(TableRowRaw(
                                Link(reference.Control.Name, appDetailsFileName),
                                Encode(reference.RuleProperty)));
                        }
                        else
                        {
                            string screenFileName = screenFileNames.GetValueOrDefault(reference.Control.Screen()?.Name, "#");
                            string controlAnchor = SanitizeAnchorId(reference.Control.Name);
                            body.Append(TableRowRaw(
                                Link(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")", screenFileName + "#" + controlAnchor),
                                Encode(reference.RuleProperty)));
                        }
                    }
                    body.AppendLine(TableEnd());
                }
            }

            SaveHtmlFile(Path.Combine(content.folderPath, variablesFileName),
                WrapInHtmlPage("Variables - " + content.Name, body.ToString(), getNavigationHtml()));
        }

        private void addAppDataSources()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.appProperties.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.appDataSources.header));
            body.AppendLine(Paragraph(content.appDataSources.infoText));

            foreach (DataSource datasource in content.appDataSources.dataSources)
            {
                if (!datasource.isSampleDataSource() || documentSampleData)
                {
                    string dsFileName = datasourceFileNames.GetValueOrDefault(datasource.Name, "#");
                    body.AppendLine(HeadingRaw(3, Link(datasource.Name, dsFileName)));

                    // Build individual datasource page
                    StringBuilder dsBody = new StringBuilder();
                    dsBody.AppendLine(Heading(1, content.appProperties.header));
                    dsBody.AppendLine(buildMetadataTable());
                    dsBody.AppendLine(Heading(3, datasource.Name));
                    dsBody.Append(TableStart("Property", "Value"));
                    dsBody.Append(TableRow("Name", datasource.Name));
                    dsBody.Append(TableRow("Type", datasource.Type));
                    dsBody.AppendLine(TableEnd());

                    dsBody.AppendLine(Heading(4, "DataSource Properties"));
                    dsBody.Append(TableStart("Property", "Value"));
                    foreach (Expression expression in datasource.Properties.OrderBy(o => o.expressionOperator))
                    {
                        if (expression.expressionOperator == "TableDefinition")
                        {
                            continue;
                        }
                        if (expression.expressionOperands.Count > 1)
                        {
                            dsBody.Append(TableRowRaw(Encode(expression.expressionOperator), AddExpressionDetails(new List<Expression> { expression })));
                        }
                        else
                        {
                            dsBody.Append(TableRow(expression.expressionOperator, (expression.expressionOperands.Count > 0) ? expression.expressionOperands[0].ToString() : ""));
                        }
                    }
                    dsBody.AppendLine(TableEnd());

                    // Render parsed Table Definition for Dataverse tables
                    Expression tableDefExpr = datasource.Properties.FirstOrDefault(o => o.expressionOperator == "TableDefinition");
                    if (tableDefExpr != null)
                    {
                        TableDefinitionInfo tdInfo = TableDefinitionHelper.Parse(tableDefExpr);
                        if (tdInfo != null)
                        {
                            dsBody.AppendLine(Heading(4, "Table Definition"));
                            dsBody.Append(TableStart("Property", "Value"));
                            foreach (var kvp in TableDefinitionHelper.GetSummaryProperties(tdInfo))
                            {
                                dsBody.Append(TableRow(kvp.Key, kvp.Value));
                            }
                            dsBody.AppendLine(TableEnd());

                            // Cross-doc link to solution table documentation
                            if (content.context?.Config?.documentSolution == true && content.context?.Solution != null
                                && !string.IsNullOrEmpty(tdInfo.LogicalName))
                            {
                                string solutionHtml = CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName);
                                string anchor = CrossDocLinkHelper.GetSolutionTableHtmlAnchor(tdInfo.LogicalName);
                                dsBody.AppendLine(Paragraph("See " + Link("full table documentation in the solution", "../" + solutionHtml + anchor)));
                            }
                        }
                    }
                    SaveHtmlFile(Path.Combine(content.folderPath, dsFileName),
                        WrapInHtmlPage("DataSource - " + datasource.Name, dsBody.ToString(), getNavigationHtml()));
                }
            }

            SaveHtmlFile(Path.Combine(content.folderPath, dataSourcesFileName),
                WrapInHtmlPage("DataSources - " + content.Name, body.ToString(), getNavigationHtml()));
        }

        private void addAppResources()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.appProperties.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.appResources.header));
            body.AppendLine(Paragraph(content.appResources.infoText));

            foreach (Resource resource in content.appResources.resources)
            {
                if (!resource.isSampleResource())
                {
                    body.AppendLine(Heading(3, resource.Name));
                    body.Append(TableStart("Property", "Value"));
                    body.Append(TableRow("Name", resource.Name));
                    body.Append(TableRow("Content", resource.Content));
                    body.Append(TableRow("Resource Kind", resource.ResourceKind));
                    if (resource.ResourceKind == "LocalFile")
                    {
                        if (content.ResourceStreams.TryGetValue(resource.Name, out MemoryStream resourceStream))
                        {
                            Expression fileName = resource.Properties.First(o => o.expressionOperator == "FileName");
                            using Stream streamToWriteTo = File.Open(content.folderPath + @"resources\" + fileName.expressionOperands[0].ToString(), FileMode.Create);
                            resourceStream.Position = 0;
                            resourceStream.CopyTo(streamToWriteTo);
                            body.Append(TableRowRaw("Resource Preview", Image(resource.Name, "resources/" + fileName.expressionOperands[0].ToString())));
                        }
                    }
                    foreach (Expression expression in resource.Properties.OrderBy(o => o.expressionOperator))
                    {
                        body.Append(TableRow(expression.expressionOperator, expression.expressionOperands?[0].ToString()));
                    }
                    body.AppendLine(TableEnd());
                }
            }

            SaveHtmlFile(Path.Combine(content.folderPath, resourcesFileName),
                WrapInHtmlPage("Resources - " + content.Name, body.ToString(), getNavigationHtml()));
        }

        private void addAppControlsOverview()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.appProperties.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.appControls.headerOverview));
            body.AppendLine(Paragraph(content.appControls.infoTextScreens));
            body.AppendLine(Paragraph(content.appControls.infoTextControls));

            foreach (ControlEntity control in content.appControls.controls.Where(o => o.Type != "appinfo"))
            {
                string screenFile = screenFileNames.GetValueOrDefault(control.Name, "#");
                body.AppendLine(HeadingRaw(3, Link("Screen: " + control.Name, screenFile)));
                body.Append(CreateControlListHtml(control));
            }

            body.AppendLine(Heading(2, content.appControls.headerScreenNavigation));
            body.AppendLine(Paragraph(content.appControls.infoTextScreenNavigation));
            body.AppendLine(ParagraphRaw(Image(content.appControls.headerScreenNavigation, content.appControls.imageScreenNavigation + ".svg")));

            SaveHtmlFile(Path.Combine(content.folderPath, controlsFileName),
                WrapInHtmlPage("Controls - " + content.Name, body.ToString(), getNavigationHtml()));
        }

        private string CreateControlListHtml(ControlEntity control)
        {
            var svgDocument = SvgDocument.FromSvg<SvgDocument>(AppControlIcons.GetControlIcon(control.Type));
            using (var bitmap = svgDocument.Draw(16, 0))
            {
                bitmap?.Save(content.folderPath + @"resources\" + control.Type + ".png");
            }
            string screenFile = screenFileNames.GetValueOrDefault(control.Screen()?.Name, "#");
            string controlAnchor = SanitizeAnchorId(control.Name);
            StringBuilder sb = new StringBuilder("<ul>");
            sb.Append("<li>");
            sb.Append($"<a href=\"{Encode(screenFile)}#{controlAnchor}\">");
            sb.Append(ImageWithClass(control.Type, "resources/" + control.Type + ".png", "icon-inline"));
            sb.Append($" {Encode(control.Name)}</a>");
            foreach (ControlEntity child in control.Children.OrderBy(o => o.Name).ToList())
            {
                sb.Append(CreateControlListHtml(child));
            }
            sb.Append("</li></ul>");
            return sb.ToString();
        }

        private void addDetailedAppControls()
        {
            foreach (ControlEntity screen in content.appControls.controls.Where(o => o.Type == "screen").OrderBy(o => o.Name).ToList())
            {
                StringBuilder body = new StringBuilder();
                body.AppendLine(Heading(1, content.appProperties.header));
                body.AppendLine(buildMetadataTable());
                body.AppendLine(HeadingWithId(2, screen.Name, SanitizeAnchorId(screen.Name)));
                body.Append(buildControlTable(screen));

                foreach (ControlEntity control in content.appControls.allControls.Where(o => o.Type != "appinfo" && o.Type != "screen" && screen.Equals(o.Screen())).OrderBy(o => o.Name).ToList())
                {
                    body.AppendLine(HeadingWithId(2, control.Name, SanitizeAnchorId(control.Name)));
                    body.Append(buildControlTable(control));
                }

                string screenFile = screenFileNames.GetValueOrDefault(screen.Name, screen.Name + ".html");
                SaveHtmlFile(Path.Combine(content.folderPath, screenFile),
                    WrapInHtmlPage("Screen: " + screen.Name + " - " + content.Name, body.ToString(), getNavigationHtml()));
            }
        }

        private string buildControlTable(ControlEntity control)
        {
            Entity defaultEntity = DefaultChangeHelper.GetEntityDefaults(control.Type);
            StringBuilder sb = new StringBuilder();
            var svgDocument = SvgDocument.FromSvg<SvgDocument>(AppControlIcons.GetControlIcon(control.Type));
            using (var bitmap = svgDocument.Draw(16, 0))
            {
                bitmap?.Save(content.folderPath + @"resources\" + control.Type + ".png");
            }
            sb.Append(TableStart("Property", "Value"));
            sb.Append(TableRowRaw(Image(control.Type, "resources/" + control.Type + ".png"), "Type: " + Encode(control.Type)));
            string category = "";

            foreach (Rule rule in control.Rules.OrderBy(o => o.Category).ThenBy(o => o.Property).ToList())
            {
                string defaultValue = defaultEntity?.Rules.Find(r => r.Property == rule.Property)?.InvariantScript;
                if (String.IsNullOrEmpty(defaultValue)) defaultValue = DefaultChangeHelper.DefaultValueIfUnknown;
                if (!documentChangedDefaultsOnly || (defaultValue != rule.InvariantScript))
                {
                    if (!content.ColourProperties.Contains(rule.Property))
                    {
                        if (rule.Category != category)
                        {
                            if (category != "")
                            {
                                sb.AppendLine(TableEnd());
                            }
                            category = rule.Category;
                            sb.AppendLine(Heading(3, category));
                            sb.Append(TableStart("Property", "Value"));
                        }
                        if (rule.InvariantScript.StartsWith("RGBA("))
                        {
                            sb.Append(buildColorRow(rule, defaultValue));
                        }
                        else
                        {
                            sb.Append(buildPropertyRow(rule, defaultValue));
                        }
                    }
                }
            }
            sb.AppendLine(TableEnd());

            // Color Properties
            bool colourPropertiesHeaderAdded = false;
            StringBuilder colourSb = new StringBuilder();
            foreach (string property in content.ColourProperties)
            {
                Rule rule = control.Rules.Find(o => o.Property == property);
                if (rule != null)
                {
                    string defaultValue = defaultEntity?.Rules.Find(r => r.Property == rule.Property)?.InvariantScript;
                    if (String.IsNullOrEmpty(defaultValue)) defaultValue = DefaultChangeHelper.DefaultValueIfUnknown;
                    if (!documentChangedDefaultsOnly || defaultValue != rule.InvariantScript)
                    {
                        if (!colourPropertiesHeaderAdded)
                        {
                            colourSb.AppendLine(Heading(3, "Color Properties"));
                            colourSb.Append(TableStart("Property", "Value"));
                            colourPropertiesHeaderAdded = true;
                        }
                        if (rule.InvariantScript.StartsWith("RGBA("))
                        {
                            colourSb.Append(buildColorRow(rule, defaultValue));
                        }
                        else
                        {
                            colourSb.Append(TableRowRaw(Encode(rule.Property), CodeBlock(rule.InvariantScript)));
                        }
                    }
                }
            }
            if (colourPropertiesHeaderAdded)
            {
                colourSb.AppendLine(TableEnd());
                sb.Append(colourSb);
            }


            // Child & Parent Controls
            if (control.Children.Count > 0 || control.Parent != null)
            {
                sb.AppendLine(Heading(3, "Child & Parent Controls"));
                sb.Append(TableStart("Property", "Value"));
                foreach (ControlEntity childControl in control.Children)
                {
                    sb.Append(TableRow("Child Control", childControl.Name));
                }
                if (control.Parent != null)
                {
                    sb.Append(TableRow("Parent Control", control.Parent.Name));
                }
                sb.AppendLine(TableEnd());
            }
            return sb.ToString();
        }

        private string buildColorRow(Rule rule, string defaultValue)
        {
            string colour = ColourHelper.ParseColor(rule.InvariantScript[..(rule.InvariantScript.IndexOf(')') + 1)]);
            StringBuilder valueHtml = new StringBuilder();
            valueHtml.Append(Encode(rule.InvariantScript));
            if (!String.IsNullOrEmpty(colour))
            {
                valueHtml.Append($" <span class=\"color-swatch\" style=\"background-color:{colour}\"></span>");
            }
            if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
            {
                valueHtml.Append("<br/><span class=\"changed-value\">" + Encode(rule.InvariantScript) + "</span>");
                valueHtml.Append(" <span class=\"default-value\">" + Encode(defaultValue) + "</span>");
            }
            return TableRowRaw(Encode(rule.Property), valueHtml.ToString());
        }

        private string buildPropertyRow(Rule rule, string defaultValue)
        {
            if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
            {
                string valueHtml = $"<span class=\"changed-value\">{CodeBlock(rule.InvariantScript)}</span> <span class=\"default-value\">{Encode(defaultValue)}</span>";
                return TableRowRaw(Encode(rule.Property), valueHtml);
            }
            return TableRowRaw(Encode(rule.Property), CodeBlock(rule.InvariantScript));
        }
    }
}
