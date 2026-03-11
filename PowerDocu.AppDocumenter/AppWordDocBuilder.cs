using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerDocu.Common;

namespace PowerDocu.AppDocumenter
{
    class AppWordDocBuilder : WordDocBuilder
    {
        private readonly AppDocumentationContent content;
        private bool documentChangedDefaultsOnly;
        private bool showDefaults;
        private bool documentSampleData;
        private string template;

        public AppWordDocBuilder(AppDocumentationContent contentDocumentation, string template, bool documentChangedDefaultsOnly = false, bool showDefaults = true, bool documentSampleData = false, bool addTableOfContents = false)
        {
            content = contentDocumentation;
            this.documentChangedDefaultsOnly = documentChangedDefaultsOnly;
            this.showDefaults = showDefaults;
            this.documentSampleData = documentSampleData;
            this.template = template;
            Directory.CreateDirectory(content.folderPath);
            // Main document: everything except detailed controls
            string filename = InitializeWordDocument(content.folderPath + content.filename, template);
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true))
            {
                mainPart = wordDocument.MainDocumentPart;
                body = mainPart.Document.Body;
                PrepareDocument(!String.IsNullOrEmpty(template));
                if (addTableOfContents) AddTableOfContents();
                addAppProperties();
                addAppVariablesInfo();
                addAppDataSources();
                addAppResources();
                addAppControlsOverview(wordDocument);
            }
            // Per-screen documents: detailed controls for each screen
            foreach (ControlEntity screen in content.appControls.controls.Where(o => o.Type == "screen").OrderBy(o => o.Name).ToList())
            {
                addScreenDocument(screen);
            }
            NotificationHelper.SendNotification("Created Word documentation for " + contentDocumentation.Name);
        }

        private void addAppProperties()
        {
            AddHeading(content.appProperties.header, "Heading1");
            body.AppendChild(new Paragraph(new Run()));
            Table table = CreateTable();
            table.Append(CreateRow(new Text("App Name"), new Text(content.Name)));
            if (content.context?.Solution != null)
            {
                table.Append(CreateRow(new Text("Solution"), new Text(content.context.Solution.UniqueName)));
            }
            //if there is a custom logo we add it to the documentation as well. Icon based logos currently not supported
            if (!String.IsNullOrEmpty(content.appProperties.appLogo))
            {
                if (content.ResourceStreams.TryGetValue(content.appProperties.appLogo, out MemoryStream resourceStream))
                {
                    Drawing icon;
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                    int imageWidth, imageHeight;
                    using (var image = Image.FromStream(resourceStream, false, false))
                    {
                        imageWidth = image.Width;
                        imageHeight = image.Height;
                    }
                    resourceStream.Position = 0;
                    imagePart.FeedData(resourceStream);
                    int usedWidth = (imageWidth > 400) ? 400 : imageWidth;
                    icon = InsertImage(mainPart.GetIdOfPart(imagePart), usedWidth, (int)(usedWidth * imageHeight / imageWidth));
                    TableRow tr = CreateRow(new Text("App Logo"), icon);
                    if (!String.IsNullOrEmpty(content.appProperties.appBackgroundColour))
                    {
                        TableCell tc = (TableCell)tr.LastChild;
                        EnsureCellProperties(tc).Append(CreateCellShading(ColourHelper.ParseColor(content.appProperties.appBackgroundColour)));
                    }
                    table.Append(tr);
                }
            }
            table.Append(CreateRow(new Text(content.appProperties.headerDocumentationGenerated), new Text(PowerDocuReleaseHelper.GetTimestampWithVersion())));
            Table statisticsTable = CreateTable();
            foreach (KeyValuePair<string, string> stats in content.appProperties.statisticsTable)
            {
                statisticsTable.Append(CreateRow(new Text(stats.Key), new Text(stats.Value)));
            }
            table.Append(CreateRow(new Text(content.appProperties.headerAppStatistics), statisticsTable));
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
            AddHeading(content.appProperties.headerAppProperties, "Heading1");
            body.AppendChild(new Paragraph(new Run()));
            table = CreateTable();
            foreach (Expression property in content.appProperties.appProperties)
            {
                if (!content.appProperties.propertiesToSkip.Contains(property.expressionOperator))
                {
                    AddExpressionTable(property, table, 1, false, true);
                }
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
            AddHeading(content.appProperties.headerAppInfo, "Heading1");
            body.AppendChild(new Paragraph(new Run()));
            addAppControlsTable(content.appControls.controls.First<ControlEntity>(o => o.Type == "appinfo"));
            AddHeading(content.appProperties.headerAppPreviewFlags, "Heading1");
            body.AppendChild(new Paragraph(new Run()));
            table = CreateTable();
            Expression appPreviewsFlagProperty = content.appProperties.appPreviewsFlagProperty;
            if (appPreviewsFlagProperty != null)
            {
                foreach (Expression flagProp in appPreviewsFlagProperty.expressionOperands)
                {
                    AddExpressionTable(flagProp, table, 1, false, true);
                }
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAppVariablesInfo()
        {
            AddHeading(content.appVariablesInfo.header, "Heading1");
            body.AppendChild(new Paragraph(new Run(new Text(content.appVariablesInfo.infoText))));
            AddHeading(content.appVariablesInfo.headerGlobalVariables, "Heading2");
            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Variable Name"), new Text("Used In")));
            foreach (string var in content.appVariablesInfo.globalVariables)
            {
                Table varReferenceTable = CreateTable();
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out List<ControlPropertyReference> references);
                if (references != null)
                {
                    varReferenceTable.Append(CreateHeaderRow(new Text("Control"), new Text("Property")));
                    foreach (ControlPropertyReference reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        if (reference.Control.Type == "appinfo")
                        {
                            varReferenceTable.Append(CreateRow(new Text("App"), new Text(reference.RuleProperty)));
                        }
                        else
                        {
                            varReferenceTable.Append(CreateRow(new Text(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")"), new Text(reference.RuleProperty)));
                        }
                    }
                }
                table.Append(CreateRow(new Text(var), varReferenceTable));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
            AddHeading(content.appVariablesInfo.headerContextVariables, "Heading2");
            table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Variable Name"), new Text("Used In")));
            foreach (string var in content.appVariablesInfo.contextVariables)
            {
                Table varReferenceTable = CreateTable();
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(var, out List<ControlPropertyReference> references);
                if (references != null)
                {
                    varReferenceTable.Append(CreateHeaderRow(new Text("Control"), new Text("Property")));
                    foreach (ControlPropertyReference reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        varReferenceTable.Append(CreateRow(new Text(reference.Control.Name + " (" + reference.Control.Screen()?.Name + ")"), new Text(reference.RuleProperty)));
                    }
                }
                table.Append(CreateRow(new Text(var), varReferenceTable));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
            AddHeading(content.appVariablesInfo.headerCollections, "Heading2");
            table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Collection Name"), new Text("Used In")));
            foreach (string coll in content.appVariablesInfo.collections)
            {
                Table collReferenceTable = CreateTable();
                content.appVariablesInfo.variableCollectionControlReferences.TryGetValue(coll, out List<ControlPropertyReference> references);
                if (references != null)
                {
                    collReferenceTable.Append(CreateHeaderRow(new Text("Control"), new Text("Property")));
                    foreach (ControlPropertyReference reference in references.OrderBy(o => o.Control.Name).ThenBy(o => o.RuleProperty))
                    {
                        collReferenceTable.Append(CreateRow(new Text(reference.Control.Name), new Text(reference.RuleProperty)));
                    }
                }
                table.Append(CreateRow(new Text(coll), collReferenceTable));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAppControlsOverview(WordprocessingDocument wordDoc)
        {
            AddHeading(content.appControls.headerOverview, "Heading1");
            body.AppendChild(new Paragraph(new Run(new Text(content.appControls.infoTextScreens))));
            body.AppendChild(new Paragraph(new Run(new Text(content.appControls.infoTextControls))));
            foreach (ControlEntity control in content.appControls.controls.Where(o => o.Type != "appinfo"))
            {
                AddHeading("Screen: " + control.Name, "Heading2");
                AppendControlTree(control, 0);
                body.AppendChild(new Paragraph(new Run(new Break())));
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
            AddHeading(content.appControls.headerScreenNavigation, "Heading2");
            body.AppendChild(new Paragraph(new Run(new Text(content.appControls.infoTextScreenNavigation))));
            ImagePart imagePart = wordDoc.MainDocumentPart.AddImagePart(ImagePartType.Png);
            int imageWidth, imageHeight;
            using (FileStream stream = new FileStream(content.folderPath + content.appControls.imageScreenNavigation + ".png", FileMode.Open))
            {
                using (var image = Image.FromStream(stream, false, false))
                {
                    imageWidth = image.Width;
                    imageHeight = image.Height;
                }
                stream.Position = 0;
                imagePart.FeedData(stream);
            }
            ImagePart svgPart = wordDoc.MainDocumentPart.AddNewPart<ImagePart>("image/svg+xml", "rId" + (new Random()).Next(100000, 999999));
            using (FileStream stream = new FileStream(content.folderPath + content.appControls.imageScreenNavigation + ".svg", FileMode.Open))
            {
                svgPart.FeedData(stream);
            }
            body.AppendChild(new Paragraph(new Run(
                InsertSvgImage(wordDoc.MainDocumentPart.GetIdOfPart(svgPart), wordDoc.MainDocumentPart.GetIdOfPart(imagePart), imageWidth, imageHeight)
            )));
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        /// <summary>
        /// Renders the control hierarchy as indented paragraphs with inline icons,
        /// instead of deeply nested tables. Each level indents by 360 twips (~0.25 inch).
        /// </summary>
        private void AppendControlTree(ControlEntity control, int depth)
        {
            const int indentPerLevel = 360; // twips per nesting level
            string controlType = control.Type;

            // Build the paragraph: [icon] ControlName [Type]
            Drawing icon = InsertSvgImage(mainPart, AppControlIcons.GetControlIcon(controlType), 16, 16);
            Paragraph para = new Paragraph();

            // Set indentation based on depth
            if (depth > 0)
            {
                para.ParagraphProperties = new ParagraphProperties(
                    new Indentation() { Left = (depth * indentPerLevel).ToString() });
            }

            // Icon run
            para.Append(new Run(icon));

            // Space between icon and text
            para.Append(new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }));

            // Control name + type 
            para.Append(new Run(new Text(control.Name + " [" + controlType + "]")));

            body.AppendChild(para);

            // Recurse into children
            foreach (ControlEntity child in control.Children.OrderBy(o => o.Name).ToList())
            {
                AppendControlTree(child, depth + 1);
            }
        }

        private void addScreenDocument(ControlEntity screen)
        {
            string screenFileName = content.folderPath + content.filename + " - " + CharsetHelper.GetSafeName(screen.Name) + " Screen";
            string filename = InitializeWordDocument(screenFileName, template);
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true);
            mainPart = wordDocument.MainDocumentPart;
            body = mainPart.Document.Body;
            PrepareDocument(!String.IsNullOrEmpty(template));
            // Metadata header
            AddHeading(content.appProperties.header, "Heading1");
            Table metaTable = CreateTable();
            metaTable.Append(CreateRow(new Text("App Name"), new Text(content.Name)));
            metaTable.Append(CreateRow(new Text(content.appProperties.headerDocumentationGenerated), new Text(PowerDocuReleaseHelper.GetTimestampWithVersion())));
            body.Append(metaTable);
            body.AppendChild(new Paragraph(new Run(new Break())));
            // Screen heading and controls
            AddHeading(screen.Name, "Heading2");
            body.AppendChild(new Paragraph(new Run()));
            addAppControlsTable(screen);
            foreach (ControlEntity control in content
                .appControls
                .allControls
                .Where(o => o.Type != "appinfo" && o.Type != "screen" && screen.Equals(o.Screen()))
                .OrderBy(o => o.Name)
                .ToList())
            {
                AddHeading(control.Name, "Heading3");
                body.AppendChild(new Paragraph(new Run()));
                addAppControlsTable(control);
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAppControlsTable(ControlEntity control)
        {
            Entity defaultEntity = DefaultChangeHelper.GetEntityDefaults(control.Type);
            Table table = CreateTable();
            Table typeTable = CreateTable(BorderValues.None);
            typeTable.Append(CreateRow(InsertSvgImage(mainPart, AppControlIcons.GetControlIcon(control.Type), 16, 16), new Text(control.Type)));
            table.Append(CreateRow(new Text("Type"), typeTable));
            string category = "";
            foreach (Rule rule in control.Rules.OrderBy(o => o.Category).ThenBy(o => o.Property).ToList())
            {
                string defaultValue = defaultEntity?.Rules.Find(r => r.Property == rule.Property)?.InvariantScript;
                if (String.IsNullOrEmpty(defaultValue))
                    defaultValue = DefaultChangeHelper.DefaultValueIfUnknown;
                if (!documentChangedDefaultsOnly || (defaultValue != rule.InvariantScript))
                {
                    if (!content.ColourProperties.Contains(rule.Property))
                    {
                        if (rule.Category != category)
                        {
                            category = rule.Category;
                            table.Append(CreateMergedRow(new Text(category), 2, WordDocBuilder.cellHeaderBackground));
                        }
                        if (rule.InvariantScript.StartsWith("RGBA("))
                        {
                            table.Append(CreateColorTable(rule, defaultValue));
                        }
                        else
                        {
                            table.Append(CreateRowForControlProperty(rule, defaultValue));
                        }
                    }
                }
            }
            bool colourPropertiesHeaderAdded = false;
            foreach (string property in content.ColourProperties)
            {
                Rule rule = control.Rules.Find(o => o.Property == property);
                if (rule != null)
                {
                    string defaultValue = defaultEntity?.Rules.Find(r => r.Property == rule.Property)?.InvariantScript;
                    if (String.IsNullOrEmpty(defaultValue))
                        defaultValue = DefaultChangeHelper.DefaultValueIfUnknown;
                    if (!documentChangedDefaultsOnly || defaultValue != rule.InvariantScript)
                    {
                        //we only need to add this once, and only if we add content
                        if (!colourPropertiesHeaderAdded)
                        {
                            table.Append(CreateMergedRow(new Text("Color Properties"), 2, WordDocBuilder.cellHeaderBackground));
                            colourPropertiesHeaderAdded = true;
                        }
                        if (rule.InvariantScript.StartsWith("RGBA("))
                        {
                            table.Append(CreateColorTable(rule, defaultValue));
                        }
                        else
                        {
                            table.Append(CreateRowForControlProperty(rule, defaultValue));
                        }
                    }
                }
            }
            if (control.Children.Count > 0 || control.Parent != null)
            {
                table.Append(CreateMergedRow(new Text("Child & Parent Controls"), 2, WordDocBuilder.cellHeaderBackground));
                Table childtable = CreateTable(BorderValues.None);
                foreach (ControlEntity childControl in control.Children)
                {
                    childtable.Append(CreateRow(new Text(childControl.Name)));
                }
                table.Append(CreateRow(new Text("Child Controls"), childtable));
                if (control.Parent != null)
                {
                    table.Append(CreateRow(new Text("Parent Control"), new Text(control.Parent.Name)));
                }
            }
            //todo isLocked property could be documented
            /* //Other properties are likely not needed for documentation, still keeping this code in case we want to show them at some point
            table.Append(CreateMergedRow(new Text("Properties"), 2, WordDocBuilder.cellHeaderBackground));
            foreach (Expression expression in control.Properties)
            {
                AddExpressionTable(expression, table);
            }*/
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private TableRow CreateRowForControlProperty(Rule rule, string defaultValue)
        {
            OpenXmlElement value = new Text(rule.InvariantScript);
            if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
            {
                value = CreateTable(BorderValues.None);
                value.Append(CreateChangedDefaultColourRow(CreateRunWithLinebreaks(rule.InvariantScript), new Text(defaultValue)));
            }
            return CreateRow(new Text(rule.Property), value);
        }

        private TableRow CreateChangedDefaultColourRow(OpenXmlElement firstColumnElement, OpenXmlElement secondColumnElement)
        {
            TableCellWidth fiftyPercentWidth = new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = "2500" };
            TableRow tr = CreateRow(firstColumnElement, secondColumnElement);
            //update the cell with the current value
            TableCell tc = (TableCell)tr.FirstChild;
            var cellProps = EnsureCellProperties(tc);
            cellProps.Append(CreateCellShading("ccffcc"));
            cellProps.TableCellWidth = (TableCellWidth)fiftyPercentWidth.Clone();
            //update the cell with the default value
            tc = (TableCell)tr.LastChild;
            cellProps = EnsureCellProperties(tc);
            cellProps.Append(CreateCellShading("ffcccc"));
            cellProps.TableCellWidth = (TableCellWidth)fiftyPercentWidth.Clone();
            return tr;
        }

        private TableRow CreateColorTable(Rule rule, string defaultValue)
        {
            Table colorTable = CreateTable(BorderValues.None);
            colorTable.Append(CreateRow(new Text(rule.InvariantScript)));
            string colour = ColourHelper.ParseColor(rule.InvariantScript[..(rule.InvariantScript.IndexOf(')') + 1)]);
            if (!String.IsNullOrEmpty(colour))
            {
                colorTable.Append(CreateMergedRow(new Text(""), 1, colour));
            }
            if (showDefaults && defaultValue != rule.InvariantScript && !content.appControls.controlPropertiesToSkip.Contains(rule.Property))
            {
                Table defaultTable = CreateTable(BorderValues.None);
                defaultTable.Append(CreateRow(new Text(defaultValue)));
                string defaultColour = ColourHelper.ParseColor(defaultValue);
                if (!String.IsNullOrEmpty(defaultColour))
                {
                    defaultTable.Append(CreateMergedRow(new Text(""), 1, defaultColour));
                }
                Table changesTable = CreateTable(BorderValues.None);
                changesTable.Append(CreateChangedDefaultColourRow(colorTable, defaultTable));
                return CreateRow(new Text(rule.Property), changesTable);
            }
            return CreateRow(new Text(rule.Property), colorTable);
        }

        private void addAppDataSources()
        {
            AddHeading(content.appDataSources.header, "Heading1");
            body.AppendChild(new Paragraph(new Run(new Text(content.appDataSources.infoText))));
            foreach (DataSource datasource in content.appDataSources.dataSources)
            {
                if (!datasource.isSampleDataSource() || documentSampleData)
                {
                    AddHeading(datasource.Name, "Heading2");
                    body.AppendChild(new Paragraph(new Run()));
                    Table table = CreateTable();
                    table.Append(CreateRow(new Text("Name"), new Text(datasource.Name)));
                    table.Append(CreateRow(new Text("Type"), new Text(datasource.Type)));
                    table.Append(CreateMergedRow(new Text("DataSource Properties"), 2, WordDocBuilder.cellHeaderBackground));
                    foreach (Expression expression in datasource.Properties.OrderBy(o => o.expressionOperator))
                    {
                        if (expression.expressionOperator == "TableDefinition")
                        {
                            AddTableDefinitionSummary(expression, table);
                        }
                        else
                        {
                            AddExpressionTable(expression, table);
                        }
                    }
                    body.Append(table);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void AddTableDefinitionSummary(Expression tableDefinition, Table table)
        {
            var info = TableDefinitionHelper.Parse(tableDefinition);
            if (info == null)
            {
                // Fallback to raw expression rendering if parsing fails
                AddExpressionTable(tableDefinition, table);
                return;
            }

            table.Append(CreateMergedRow(new Text("Table Definition"), 2, WordDocBuilder.cellHeaderBackground));
            foreach (var prop in TableDefinitionHelper.GetSummaryProperties(info))
            {
                table.Append(CreateRow(new Text(prop.Key), new Text(prop.Value)));
            }

            if (info.Privileges.Count > 0)
            {
                table.Append(CreateMergedRow(new Text("Privileges"), 2, WordDocBuilder.cellHeaderBackground));
                foreach (var priv in info.Privileges)
                {
                    table.Append(CreateRow(new Text(priv.Name ?? ""), new Text(priv.PrivilegeType ?? "")));
                }
            }
        }

        private void addAppResources()
        {
            AddHeading(content.appResources.header, "Heading1");
            body.AppendChild(new Paragraph(new Run(new Text(content.appResources.infoText))));
            foreach (Resource resource in content.appResources.resources)
            {
                if (!resource.isSampleResource())
                {
                    AddHeading(resource.Name, "Heading2");
                    body.AppendChild(new Paragraph(new Run()));
                    Table table = CreateTable();
                    table.Append(CreateRow(new Text("Name"), new Text(resource.Name)));
                    table.Append(CreateRow(new Text("Content"), new Text(resource.Content)));
                    table.Append(CreateRow(new Text("Resource Kind"), new Text(resource.ResourceKind)));
                    if (resource.ResourceKind == "LocalFile" && content.ResourceStreams.TryGetValue(resource.Name, out MemoryStream resourceStream))
                    {
                        try
                        {
                            Drawing icon = null;
                            Expression fileName = resource.Properties.First(o => o.expressionOperator == "FileName");
                            if (fileName.expressionOperands[0].ToString().EndsWith("svg", StringComparison.OrdinalIgnoreCase))
                            {
                                string svg = Encoding.Default.GetString(resourceStream.ToArray());
                                icon = InsertSvgImage(mainPart, svg, 400, 400);
                            }
                            else
                            {
                                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                                int imageWidth, imageHeight;
                                using (var image = Image.FromStream(resourceStream, false, false))
                                {
                                    imageWidth = image.Width;
                                    imageHeight = image.Height;
                                }
                                resourceStream.Position = 0;
                                imagePart.FeedData(resourceStream);
                                int usedWidth = (imageWidth > 400) ? 400 : imageWidth;
                                icon = InsertImage(mainPart.GetIdOfPart(imagePart), usedWidth, (int)(usedWidth * imageHeight / imageWidth));
                            }
                            table.Append(CreateRow(new Text("Resource Preview"), icon));
                        }
                        catch (Exception e)
                        {
                            table.Append(CreateRow(new Text("Resource Preview"), new Text("Resource Preview is not available, media file is invalid.")));
                        }
                    }
                    table.Append(CreateMergedRow(new Text("Resource Properties"), 2, WordDocBuilder.cellHeaderBackground));
                    foreach (Expression expression in resource.Properties.OrderBy(o => o.expressionOperator))
                    {
                        AddExpressionTable(expression, table);
                    }
                    body.Append(table);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
        }
    }
}