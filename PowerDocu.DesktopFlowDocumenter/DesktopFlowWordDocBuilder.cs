using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerDocu.Common;

namespace PowerDocu.DesktopFlowDocumenter
{
    class DesktopFlowWordDocBuilder : WordDocBuilder
    {
        private readonly DesktopFlowDocumentationContent content;

        public DesktopFlowWordDocBuilder(DesktopFlowDocumentationContent contentDocumentation, string template)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            string filename = InitializeWordDocument(content.folderPath + content.filename, template);
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true))
            {
                mainPart = wordDocument.MainDocumentPart;
                body = mainPart.Document.Body;
                PrepareDocument(!String.IsNullOrEmpty(template));

                addOverview();
                addFlowDiagram();
                addActionSteps();
                addVariables();
                if (content.flow.ControlFlowBlocks.Count > 0)
                    addControlFlow();
                if (content.flow.Subflows.Count > 0)
                    addSubflows();
                addModules();
                if (content.flow.Connectors.Count > 0)
                    addConnectors();
                if (content.flow.EnvironmentVariables.Count > 0)
                    addEnvironmentVariables();
                addProperties();
            }
            NotificationHelper.SendNotification("Created Word documentation for Desktop Flow: " + content.flow.GetDisplayName());
        }

        private void addOverview()
        {
            AddHeading(content.flow.GetDisplayName(), "Heading1");
            body.AppendChild(new Paragraph(new Run()));

            Table table = CreateTable();
            table.Append(CreateRow(new Text("Name"), new Text(content.flow.GetDisplayName())));
            table.Append(CreateRow(new Text("State"), new Text(content.flow.GetStateLabel())));
            table.Append(CreateRow(new Text("Action Steps"), new Text(content.flow.ActionSteps.Count.ToString())));
            table.Append(CreateRow(new Text("Variables"), new Text(content.flow.Variables.Count.ToString())));
            table.Append(CreateRow(new Text("Modules Used"), new Text(content.flow.Modules.Count.ToString())));
            if (content.flow.Subflows.Count > 0)
                table.Append(CreateRow(new Text("Subflows"), new Text(content.flow.Subflows.Count.ToString())));
            if (!string.IsNullOrEmpty(content.flow.GetEngineVersionString()))
                table.Append(CreateRow(new Text("Engine Version"), new Text(content.flow.GetEngineVersionString())));
            if (content.flow.PowerFxEnabled)
                table.Append(CreateRow(new Text("Power Fx"), new Text("Enabled" + (!string.IsNullOrEmpty(content.flow.PowerFxVersion) ? $" (v{content.flow.PowerFxVersion})" : ""))));
            if (!string.IsNullOrEmpty(content.flow.Description))
                table.Append(CreateRow(new Text("Description"), new Text(content.flow.Description)));
            table.Append(CreateRow(new Text(content.headerDocumentationGenerated),
                new Text(PowerDocuReleaseHelper.GetTimestampWithVersion())));
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addFlowDiagram()
        {
            string pngFile = Path.Combine(content.folderPath, "desktopflow-detailed.png");
            if (!File.Exists(pngFile)) return;

            AddHeading("Flow Diagram", "Heading2");
            try
            {
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                int imageWidth, imageHeight;
                using (FileStream stream = new FileStream(pngFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    using (var image = Image.FromStream(stream, false, false))
                    {
                        imageWidth = image.Width;
                        imageHeight = image.Height;
                    }
                    stream.Position = 0;
                    imagePart.FeedData(stream);
                }
                int usedWidth = (imageWidth > 600) ? 600 : imageWidth;
                DocumentFormat.OpenXml.Wordprocessing.Drawing drawing = InsertImage(mainPart.GetIdOfPart(imagePart), usedWidth, (int)(usedWidth * imageHeight / imageWidth));
                body.AppendChild(new Paragraph(new Run(drawing)));
            }
            catch { }
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addActionSteps()
        {
            AddHeading(content.headerActionSteps, "Heading2");

            if (content.flow.ActionSteps.Count == 0)
            {
                body.AppendChild(new Paragraph(new Run(new Text("No action steps found."))));
                return;
            }

            body.AppendChild(new Paragraph(new Run(
                new Text($"This Desktop Flow has {content.flow.ActionSteps.Count} action step(s)."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("#"), new Text("Module"), new Text("Action"), new Text("Parameters"), new Text("Output Variables")));

            foreach (var step in content.flow.ActionSteps.OrderBy(s => s.Order))
            {
                string parameters = step.Parameters.Count > 0
                    ? string.Join(", ", step.Parameters.Select(p => $"{p.Key}: {p.Value}"))
                    : "";
                string outputs = step.OutputVariables.Count > 0
                    ? string.Join(", ", step.OutputVariables)
                    : "";

                table.Append(CreateRow(
                    new Text(step.Order.ToString()),
                    new Text(step.ModuleName),
                    new Text(step.FullActionName),
                    new Text(parameters),
                    new Text(outputs)
                ));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addVariables()
        {
            AddHeading(content.headerVariables, "Heading2");

            if (content.flow.Variables.Count == 0)
            {
                body.AppendChild(new Paragraph(new Run(new Text("No variables found."))));
                return;
            }

            body.AppendChild(new Paragraph(new Run(
                new Text($"This Desktop Flow has {content.flow.Variables.Count} variable(s)."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Name"), new Text("Type"), new Text("Input"), new Text("Output"), new Text("Sensitive"), new Text("Initial Value")));

            foreach (var variable in content.flow.Variables.OrderBy(v => v.Name))
            {
                table.Append(CreateRow(
                    new Text(variable.Name),
                    new Text(variable.Type ?? ""),
                    new Text(variable.IsInput ? "Yes" : ""),
                    new Text(variable.IsOutput ? "Yes" : ""),
                    new Text(variable.IsSensitive ? "Yes" : ""),
                    new Text(variable.InitialValue ?? "")
                ));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addControlFlow()
        {
            AddHeading(content.headerControlFlow, "Heading2");
            body.AppendChild(new Paragraph(new Run(
                new Text($"This Desktop Flow has {content.flow.ControlFlowBlocks.Count} control flow block(s)."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Type"), new Text("Condition"), new Text("Line"), new Text("Nesting Level")));

            foreach (var block in content.flow.ControlFlowBlocks)
            {
                table.Append(CreateRow(
                    new Text(block.Type),
                    new Text(block.Condition ?? ""),
                    new Text(block.StartLine.ToString()),
                    new Text(block.NestingLevel.ToString())
                ));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addSubflows()
        {
            AddHeading(content.headerSubflows, "Heading2");
            body.AppendChild(new Paragraph(new Run(
                new Text($"This Desktop Flow has {content.flow.Subflows.Count} subflow(s)."))));
            body.AppendChild(new Paragraph(new Run(new Break())));

            foreach (var subflow in content.flow.Subflows)
            {
                string subflowTitle = subflow.Name + (subflow.IsGlobal ? " (Global)" : "");
                AddHeading(subflowTitle, "Heading3");

                // Subflow flow diagram
                string pngFile = Path.Combine(content.folderPath, "desktopflow-subflow-" + CharsetHelper.GetSafeName(subflow.Name) + "-detailed.png");
                if (File.Exists(pngFile))
                {
                    try
                    {
                        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                        int imageWidth, imageHeight;
                        using (FileStream stream = new FileStream(pngFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            using (var image = Image.FromStream(stream, false, false))
                            {
                                imageWidth = image.Width;
                                imageHeight = image.Height;
                            }
                            stream.Position = 0;
                            imagePart.FeedData(stream);
                        }
                        int usedWidth = (imageWidth > 600) ? 600 : imageWidth;
                        DocumentFormat.OpenXml.Wordprocessing.Drawing drawing = InsertImage(mainPart.GetIdOfPart(imagePart), usedWidth, (int)(usedWidth * imageHeight / imageWidth));
                        body.AppendChild(new Paragraph(new Run(drawing)));
                    }
                    catch { }
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Subflow action steps
                if (subflow.ActionSteps.Count > 0)
                {
                    body.AppendChild(new Paragraph(new Run(
                        new Text($"{subflow.ActionSteps.Count} action step(s)."))));

                    Table stepTable = CreateTable();
                    stepTable.Append(CreateHeaderRow(new Text("#"), new Text("Module"), new Text("Action"), new Text("Parameters"), new Text("Output Variables")));
                    foreach (var step in subflow.ActionSteps.OrderBy(s => s.Order))
                    {
                        string parameters = step.Parameters.Count > 0
                            ? string.Join(", ", step.Parameters.Select(p => $"{p.Key}: {p.Value}"))
                            : "";
                        string outputs = step.OutputVariables.Count > 0
                            ? string.Join(", ", step.OutputVariables)
                            : "";
                        stepTable.Append(CreateRow(
                            new Text(step.Order.ToString()),
                            new Text(step.ModuleName),
                            new Text(step.FullActionName),
                            new Text(parameters),
                            new Text(outputs)
                        ));
                    }
                    body.Append(stepTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Subflow variables
                if (subflow.Variables.Count > 0)
                {
                    body.AppendChild(new Paragraph(new Run(
                        new Text($"{subflow.Variables.Count} variable(s)."))));

                    Table varTable = CreateTable();
                    varTable.Append(CreateHeaderRow(new Text("Name"), new Text("Type"), new Text("Input"), new Text("Output"), new Text("Sensitive"), new Text("Initial Value")));
                    foreach (var variable in subflow.Variables.OrderBy(v => v.Name))
                    {
                        varTable.Append(CreateRow(
                            new Text(variable.Name),
                            new Text(variable.Type ?? ""),
                            new Text(variable.IsInput ? "Yes" : ""),
                            new Text(variable.IsOutput ? "Yes" : ""),
                            new Text(variable.IsSensitive ? "Yes" : ""),
                            new Text(variable.InitialValue ?? "")
                        ));
                    }
                    body.Append(varTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Subflow control flow
                if (subflow.ControlFlowBlocks.Count > 0)
                {
                    body.AppendChild(new Paragraph(new Run(
                        new Text($"{subflow.ControlFlowBlocks.Count} control flow block(s)."))));

                    Table cfTable = CreateTable();
                    cfTable.Append(CreateHeaderRow(new Text("Type"), new Text("Condition"), new Text("Line"), new Text("Nesting Level")));
                    foreach (var block in subflow.ControlFlowBlocks)
                    {
                        cfTable.Append(CreateRow(
                            new Text(block.Type),
                            new Text(block.Condition ?? ""),
                            new Text(block.StartLine.ToString()),
                            new Text(block.NestingLevel.ToString())
                        ));
                    }
                    body.Append(cfTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
        }

        private void addModules()
        {
            AddHeading(content.headerModules, "Heading2");

            if (content.flow.Modules.Count == 0)
            {
                body.AppendChild(new Paragraph(new Run(new Text("No modules referenced."))));
                return;
            }

            body.AppendChild(new Paragraph(new Run(
                new Text($"This Desktop Flow references {content.flow.Modules.Count} module(s)."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Module"), new Text("Assembly"), new Text("Version"), new Text("Actions Used")));

            foreach (var module in content.flow.Modules.OrderBy(m => m.Name))
            {
                int actionCount = content.flow.ActionSteps.Count(a => a.ModuleName == module.Name);
                table.Append(CreateRow(
                    new Text(module.Name),
                    new Text(module.AssemblyName ?? ""),
                    new Text(module.Version ?? ""),
                    new Text(actionCount.ToString())
                ));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addConnectors()
        {
            AddHeading(content.headerConnectors, "Heading2");
            body.AppendChild(new Paragraph(new Run(
                new Text($"This Desktop Flow uses {content.flow.Connectors.Count} cloud connector(s)."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Connector ID"), new Text("Name"), new Text("Title")));

            foreach (var connector in content.flow.Connectors)
            {
                table.Append(CreateRow(
                    new Text(connector.ConnectorId ?? ""),
                    new Text(connector.Name ?? ""),
                    new Text(connector.Title ?? "")
                ));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addEnvironmentVariables()
        {
            AddHeading(content.headerEnvironmentVariables, "Heading2");

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Name"), new Text("Type"), new Text("Value")));

            foreach (var envVar in content.flow.EnvironmentVariables)
            {
                table.Append(CreateRow(
                    new Text(envVar.Name ?? ""),
                    new Text(envVar.Type ?? ""),
                    new Text(envVar.Value ?? "")
                ));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addProperties()
        {
            AddHeading(content.headerProperties, "Heading2");

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Property"), new Text("Value")));
            table.Append(CreateRow(new Text("ID"), new Text(content.flow.ID ?? "")));
            table.Append(CreateRow(new Text("Schema Version"), new Text(content.flow.SchemaVersion ?? "")));
            table.Append(CreateRow(new Text("UI Flow Type"), new Text(content.flow.UIFlowType.ToString())));
            table.Append(CreateRow(new Text("State"), new Text(content.flow.GetStateLabel())));
            table.Append(CreateRow(new Text("Is Customizable"), new Text(content.flow.IsCustomizable ? "Yes" : "No")));
            if (!string.IsNullOrEmpty(content.flow.IntroducedVersion))
                table.Append(CreateRow(new Text("Introduced Version"), new Text(content.flow.IntroducedVersion)));
            if (!string.IsNullOrEmpty(content.flow.GetEngineVersionString()))
                table.Append(CreateRow(new Text("Engine Version"), new Text(content.flow.GetEngineVersionString())));
            if (content.flow.PowerFxEnabled)
                table.Append(CreateRow(new Text("Power Fx Version"), new Text(content.flow.PowerFxVersion ?? "")));
            body.Append(table);
        }

    }
}
