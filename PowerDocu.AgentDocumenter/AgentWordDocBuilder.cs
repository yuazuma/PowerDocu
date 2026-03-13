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

namespace PowerDocu.AgentDocumenter
{
    class AgentWordDocBuilder : WordDocBuilder
    {
        private readonly AgentDocumentationContent content;

        public AgentWordDocBuilder(AgentDocumentationContent contentDocumentation, string template)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            string filename = InitializeWordDocument(content.folderPath + content.filename, template);
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true);
            mainPart = wordDocument.MainDocumentPart;
            body = mainPart.Document.Body;
            PrepareDocument(!String.IsNullOrEmpty(template));
            addAgentOverview();
            addAgentKnowledge();
            addAgentTools();
            addAgentEntities();
            addAgentVariables();
            addAgentChannels();
            addAgentSettings();
            addAgentTopicsOverview();
            addAgentTopicDetails();
            NotificationHelper.SendNotification("Created Word documentation for " + contentDocumentation.filename);
        }

        private void addAgentOverview()
        {
            AddHeading("Agent - " + content.filename, "Heading1");

            Table table = CreateTable();
            table.Append(CreateRow(new Text("Agent Name"), new Text(content.agent.Name)));
            if (content.context?.Solution != null)
            {
                table.Append(CreateRow(new Text("Solution"), new Text(content.context.Solution.UniqueName)));
            }
            if (!String.IsNullOrEmpty(content.agent.IconBase64))
            {
                try
                {
                    Bitmap agentLogo = ImageHelper.ConvertBase64ToBitmap(content.agent.IconBase64);
                    string logoPath = content.folderPath + $"agentlogo-{content.filename.Replace(" ", "-")}.png";
                    agentLogo.Save(logoPath);
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    using (FileStream stream = new FileStream(logoPath, FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }
                    Drawing icon = InsertImage(mainPart.GetIdOfPart(imagePart), 64, 64);
                    table.Append(CreateRow(new Text("Agent Logo"), icon));
                    agentLogo.Dispose();
                }
                catch { }
            }
            table.Append(CreateRow(new Text(content.headerDocumentationGenerated), new Text(PowerDocuReleaseHelper.GetTimestampWithVersion())));
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Details section
            AddHeading(content.Details, "Heading2");

            // Description
            AddHeading(content.Description, "Heading3");
            body.AppendChild(new Paragraph(CreateRunWithLinebreaks(content.agent.GetDescription() ?? "")));

            // Orchestration
            AddHeading(content.Orchestration, "Heading3");
            body.AppendChild(new Paragraph(new Run(new Text($"{content.OrchestrationText} - {content.agent.GetOrchestration()}"))));

            // Response Model
            AddHeading(content.ResponseModel, "Heading3");
            body.AppendChild(new Paragraph(new Run(new Text(content.agent.GetResponseModel()))));

            // Instructions
            AddHeading(content.Instructions, "Heading3");
            body.AppendChild(new Paragraph(CreateRunWithLinebreaks(content.agent.GetInstructions() ?? "")));

            // Knowledge
            AddHeading(content.Knowledge, "Heading3");
            var knowledgeSources = content.agent.GetKnowledge();
            var fileKnowledge = content.agent.GetFileKnowledge();
            if (knowledgeSources.Count > 0 || fileKnowledge.Count > 0)
            {
                Table knowledgeTable = CreateTable();
                knowledgeTable.Append(CreateHeaderRow(new Text("Name"), new Text("Source Type"), new Text("Details")));
                foreach (BotComponent ks in knowledgeSources)
                {
                    string details = content.agent.GetKnowledgeDetailsSummary(ks);
                    string site = ks.GetKnowledgeSourceSite();
                    OpenXmlElement detailsCell;
                    if (!string.IsNullOrEmpty(site))
                    {
                        try
                        {
                            var rel = mainPart.AddHyperlinkRelationship(new Uri(site), true);
                            detailsCell = new Hyperlink(
                                new Run(
                                    new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = "0563C1", ThemeColor = ThemeColorValues.Hyperlink }),
                                    new Text(details)))
                            { History = OnOffValue.FromBoolean(true), Id = rel.Id };
                        }
                        catch (UriFormatException)
                        {
                            detailsCell = new Text(details);
                        }
                    }
                    else
                    {
                        detailsCell = new Text(details);
                    }
                    knowledgeTable.Append(CreateRow(new Text(ks.Name), new Text(ks.GetSourceKindDisplayName()), detailsCell));
                }
                foreach (BotComponent fk in fileKnowledge)
                {
                    string mimeType = !string.IsNullOrEmpty(fk.FileDataMimeType) ? $" ({fk.FileDataMimeType})" : "";
                    knowledgeTable.Append(CreateRow(new Text(fk.Name), new Text("File" + mimeType), new Text(fk.FileDataName ?? "")));
                }
                body.Append(knowledgeTable);
            }
            else
            {
                body.AppendChild(new Paragraph(new Run(new Text("No knowledge sources configured."))));
            }

            // Tools
            AddHeading(content.Tools, "Heading3");
            var overviewTools = content.GetResolvedToolInfos();
            if (overviewTools.Count > 0)
            {
                foreach (AgentToolInfo tool in overviewTools)
                {
                    body.AppendChild(new Paragraph(new Run(new Text($"{tool.Name} ({tool.ToolType})"))));
                }
            }
            else
            {
                body.AppendChild(new Paragraph(new Run(new Text("No tools configured."))));
            }

            // Triggers
            AddHeading(content.Triggers, "Heading3");
            var triggers = content.agent.GetTriggers();
            if (triggers.Count > 0)
            {
                foreach (BotComponent trigger in triggers.OrderBy(t => t.Name))
                {
                    var (triggerKind, flowId, connectionType) = trigger.GetTriggerDetails();
                    string triggerInfo = !string.IsNullOrEmpty(connectionType) ? $" ({connectionType})" : "";
                    body.AppendChild(new Paragraph(new Run(new Text(trigger.Name + triggerInfo))));
                }
            }
            else
            {
                body.AppendChild(new Paragraph(new Run(new Text("No triggers configured."))));
            }

            // Agents
            AddHeading(content.Agents, "Heading3");
            var overviewAgents = content.GetResolvedConnectedAgentInfos();
            if (overviewAgents.Count > 0)
            {
                Table agentsTable = CreateTable();
                agentsTable.Append(CreateHeaderRow(new Text("Name"), new Text("Connection Type"), new Text("Description"), new Text("History Type")));
                foreach (var agentInfo in overviewAgents.OrderBy(a => a.Name))
                {
                    agentsTable.Append(CreateRow(new Text(agentInfo.Name), new Text(agentInfo.ConnectionType), new Text(agentInfo.Description), new Text(agentInfo.HistoryType)));
                }
                body.AppendChild(agentsTable);
            }
            else
            {
                body.AppendChild(new Paragraph(new Run(new Text("No connected agents configured."))));
            }

            // Topics
            AddHeading(content.Topics, "Heading3");
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name))
            {
                body.AppendChild(new Paragraph(new Run(new Text(topic.Name))));
            }

            // Suggested Prompts
            AddHeading(content.SuggestedPrompts, "Heading3");
            body.AppendChild(new Paragraph(new Run(new Text(content.SuggestedPromptsText))));
            Dictionary<string, string> conversationStarters = content.agent.GetSuggestedPrompts();
            if (conversationStarters.Count > 0)
            {
                Table promptsTable = CreateTable();
                promptsTable.Append(CreateHeaderRow(new Text("Prompt Title"), new Text("Prompt")));
                foreach (var kvp in conversationStarters.OrderBy(x => x.Key))
                {
                    promptsTable.Append(CreateRow(new Text(kvp.Key), new Text(kvp.Value)));
                }
                body.Append(promptsTable);
            }
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAgentKnowledge()
        {
            var knowledgeSources = content.agent.GetKnowledge();
            var fileKnowledge = content.agent.GetFileKnowledge();
            if (knowledgeSources.Count == 0 && fileKnowledge.Count == 0) return;

            AddHeading(content.Knowledge, "Heading1");

            // Overview table
            Table overviewTable = CreateTable();
            overviewTable.Append(CreateHeaderRow(new Text("Name"), new Text("Source Type"), new Text("Official Source"), new Text("Details"), new Text("Description")));
            foreach (BotComponent ks in knowledgeSources.OrderBy(k => k.Name))
            {
                string details = content.agent.GetKnowledgeDetailsSummary(ks);
                string officialSource = ks.GetOfficialSourceDisplayName();
                string descriptionPreview = !string.IsNullOrEmpty(ks.Description) && ks.Description.Length > 100
                    ? ks.Description.Substring(0, 100) + "..."
                    : ks.Description ?? "";
                string site = ks.GetKnowledgeSourceSite();
                OpenXmlElement detailsCell;
                if (!string.IsNullOrEmpty(site))
                {
                    try
                    {
                        var rel = mainPart.AddHyperlinkRelationship(new Uri(site), true);
                        detailsCell = new Hyperlink(
                            new Run(
                                new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = "0563C1", ThemeColor = ThemeColorValues.Hyperlink }),
                                new Text(details)))
                        { History = OnOffValue.FromBoolean(true), Id = rel.Id };
                    }
                    catch (UriFormatException)
                    {
                        detailsCell = new Text(details);
                    }
                }
                else
                {
                    detailsCell = new Text(details);
                }
                overviewTable.Append(CreateRow(new Text(ks.Name), new Text(ks.GetSourceKindDisplayName()), new Text(officialSource), detailsCell, new Text(descriptionPreview)));
            }
            foreach (BotComponent fk in fileKnowledge.OrderBy(k => k.Name))
            {
                string mimeType = !string.IsNullOrEmpty(fk.FileDataMimeType) ? $" ({fk.FileDataMimeType})" : "";
                string descriptionPreview = !string.IsNullOrEmpty(fk.Description) && fk.Description.Length > 100
                    ? fk.Description.Substring(0, 100) + "..."
                    : fk.Description ?? "";
                overviewTable.Append(CreateRow(new Text(fk.Name), new Text("File" + mimeType), new Text(""), new Text(fk.FileDataName ?? ""), new Text(descriptionPreview)));
            }
            body.Append(overviewTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Detail per knowledge source
            foreach (BotComponent ks in knowledgeSources.OrderBy(k => k.Name))
            {
                addKnowledgeSourceDetail(ks);
            }
            foreach (BotComponent fk in fileKnowledge.OrderBy(k => k.Name))
            {
                addKnowledgeSourceDetail(fk);
            }
        }

        private void addKnowledgeSourceDetail(BotComponent knowledge)
        {
            AddHeading(knowledge.Name, "Heading2");

            // Properties table
            Table table = CreateTable();
            table.Append(CreateRow(new Text("Name"), new Text(knowledge.Name)));
            if (knowledge.ComponentType == 16)
            {
                table.Append(CreateRow(new Text("Source Type"), new Text(knowledge.GetSourceKindDisplayName())));
                string officialSource = knowledge.GetOfficialSourceDisplayName();
                if (!string.IsNullOrEmpty(officialSource))
                    table.Append(CreateRow(new Text("Official Source"), new Text(officialSource)));
                string site = knowledge.GetKnowledgeSourceSite();
                if (!string.IsNullOrEmpty(site))
                {
                    try
                    {
                        var rel = mainPart.AddHyperlinkRelationship(new Uri(site), true);
                        var hyperlink = new Hyperlink(
                            new Run(
                                new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = "0563C1", ThemeColor = ThemeColorValues.Hyperlink }),
                                new Text(site)))
                        { History = OnOffValue.FromBoolean(true), Id = rel.Id };
                        table.Append(CreateRow(new Text("URL"), hyperlink));
                    }
                    catch (UriFormatException)
                    {
                        table.Append(CreateRow(new Text("URL"), new Text(site)));
                    }
                }
            }
            else if (knowledge.ComponentType == 14)
            {
                string mimeType = !string.IsNullOrEmpty(knowledge.FileDataMimeType) ? $" ({knowledge.FileDataMimeType})" : "";
                table.Append(CreateRow(new Text("Source Type"), new Text("File" + mimeType)));
                if (!string.IsNullOrEmpty(knowledge.FileDataName))
                    table.Append(CreateRow(new Text("File Name"), new Text(knowledge.FileDataName)));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Description
            if (!string.IsNullOrEmpty(knowledge.Description))
            {
                AddHeading("Description", "Heading3");
                body.AppendChild(new Paragraph(CreateRunWithLinebreaks(knowledge.Description)));
            }

            // Dataverse-specific: Selected Tables and Synonyms
            if (knowledge.ComponentType == 16 && knowledge.GetSourceKindDisplayName() == "Dataverse")
            {
                var tables = content.agent.GetDataverseTablesForKnowledge(knowledge);
                if (tables.Count > 0)
                {
                    AddHeading("Selected Tables", "Heading3");

                    Table tablesTable = CreateTable();
                    tablesTable.Append(CreateHeaderRow(new Text("Table Name"), new Text("Logical Name")));
                    foreach (var dvTable in tables.OrderBy(t => t.Name))
                    {
                        tablesTable.Append(CreateRow(new Text(dvTable.Name), new Text(dvTable.EntityLogicalName)));
                    }
                    body.Append(tablesTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));

                    // Synonyms/Glossary per table
                    foreach (var dvTable in tables.OrderBy(t => t.Name))
                    {
                        var synonyms = content.agent.GetSynonymsForEntity(dvTable);
                        if (synonyms.Count > 0)
                        {
                            AddHeading($"Glossary: {dvTable.Name}", "Heading3");

                            Table synTable = CreateTable();
                            synTable.Append(CreateHeaderRow(new Text("Column"), new Text("Description")));
                            foreach (var syn in synonyms.OrderBy(s => s.ColumnLogicalName))
                            {
                                synTable.Append(CreateRow(new Text(syn.ColumnLogicalName), new Text(syn.Description ?? "")));
                            }
                            body.Append(synTable);
                            body.AppendChild(new Paragraph(new Run(new Break())));
                        }
                    }
                }
            }
        }

        private void addAgentTopicsOverview()
        {
            AddHeading(content.Topics, "Heading1");

            // Topic Data Flow diagram
            string dataFlowPng = Path.Combine(content.folderPath, "topic-dataflow.png");
            if (File.Exists(dataFlowPng))
            {
                AddHeading("Topic Data Flow", "Heading2");
                body.AppendChild(new Paragraph(new Run(new Text("Visualization of how topics call each other via BeginDialog/ReplaceDialog, with data passed between them."))));

                try
                {
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    int imageWidth, imageHeight;
                    using (FileStream stream = new FileStream(dataFlowPng, FileMode.Open))
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
                    Drawing drawing = InsertImage(mainPart.GetIdOfPart(imagePart), usedWidth, (int)(usedWidth * imageHeight / imageWidth));
                    body.AppendChild(new Paragraph(new Run(drawing)));
                }
                catch { }
                body.AppendChild(new Paragraph(new Run(new Break())));
                body.AppendChild(new Paragraph(new Run(new Break())));

                // Summary table
                var calls = content.GetTopicDataFlowInfo();
                if (calls.Count > 0)
                {
                    Table callTable = CreateTable();
                    callTable.Append(CreateHeaderRow(new Text("Source Topic"), new Text("Target Topic"), new Text("Call Type"), new Text("Data Passed")));
                    foreach (var call in calls)
                    {
                        string dataPassed = call.InputBindings.Count > 0 ? string.Join(", ", call.InputBindings.Keys) : "";
                        callTable.Append(CreateRow(new Text(call.SourceTopicName), new Text(call.TargetTopicName), new Text(call.CallKind), new Text(dataPassed)));
                    }
                    body.Append(callTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }

            AddHeading("Topic List", "Heading2");
            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Name"), new Text("Type"), new Text("Trigger"), new Text("Kind")));
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                string topicType = topic.GetComponentTypeDisplayName();
                string triggerType = topic.GetTriggerTypeForTopic();
                string topicKind = topic.GetTopicKind() == "KnowledgeSourceConfiguration" ? "Knowledge" : topic.GetTopicKind();
                table.Append(CreateRow(new Text(topic.Name), new Text(topicType), new Text(triggerType), new Text(topicKind)));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAgentTopicDetails()
        {
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                AddHeading("Topic: " + topic.Name, "Heading2");

                // Metadata table
                Table table = CreateTable();
                table.Append(CreateRow(new Text("Name"), new Text(topic.Name)));
                table.Append(CreateRow(new Text("Type"), new Text(topic.GetComponentTypeDisplayName())));
                table.Append(CreateRow(new Text("Trigger"), new Text(topic.GetTriggerTypeForTopic())));
                table.Append(CreateRow(new Text("Topic Kind"), new Text(topic.GetTopicKind())));
                if (!string.IsNullOrEmpty(topic.Description))
                {
                    table.Append(CreateRow(new Text("Description"), new Text(topic.Description)));
                }
                string modelDesc = topic.GetModelDescription();
                if (!string.IsNullOrEmpty(modelDesc))
                {
                    table.Append(CreateRow(new Text("Model Description"), new Text(modelDesc)));
                }
                string startBehavior = topic.GetStartBehavior();
                if (!string.IsNullOrEmpty(startBehavior))
                {
                    table.Append(CreateRow(new Text("Start Behavior"), new Text(startBehavior)));
                }
                body.Append(table);
                body.AppendChild(new Paragraph(new Run(new Break())));

                // Trigger queries
                List<string> triggerQueries = topic.GetTriggerQueries();
                if (triggerQueries.Count > 0)
                {
                    AddHeading("Trigger Queries", "Heading3");

                    Table triggerTable = CreateTable();
                    triggerTable.Append(CreateHeaderRow(new Text("Query")));
                    foreach (string query in triggerQueries)
                    {
                        triggerTable.Append(CreateRow(new Text(query)));
                    }
                    body.Append(triggerTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Knowledge source details
                if (topic.GetTopicKind() == "KnowledgeSourceConfiguration")
                {
                    var (sourceKind, skillConfig) = topic.GetKnowledgeSourceDetails();
                    AddHeading("Knowledge Source", "Heading3");

                    Table ksTable = CreateTable();
                    if (!string.IsNullOrEmpty(sourceKind))
                        ksTable.Append(CreateRow(new Text("Source Kind"), new Text(sourceKind)));
                    if (!string.IsNullOrEmpty(skillConfig))
                        ksTable.Append(CreateRow(new Text("Skill Configuration"), new Text(skillConfig)));
                    body.Append(ksTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Variables (with Scope column)
                var variables = topic.GetTopicVariables();
                if (variables.Count > 0)
                {
                    AddHeading("Variables", "Heading3");

                    Table varTable = CreateTable();
                    varTable.Append(CreateHeaderRow(new Text("Variable"), new Text("Scope"), new Text("Context")));
                    foreach (var (variable, context) in variables)
                    {
                        string scope = variable.StartsWith("Global.") ? "Global"
                            : variable.StartsWith("System.") ? "System"
                            : "Topic";
                        varTable.Append(CreateRow(new Text(variable), new Text(scope), new Text(context)));
                    }
                    body.Append(varTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Topic flow diagram
                string graphFile = topic.getTopicFileName() + "-detailed.png";
                string graphFilePath = Path.Combine(content.folderPath, "Topics", graphFile);
                if (File.Exists(graphFilePath))
                {
                    AddHeading("Topic Flow", "Heading3");

                    try
                    {
                        ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                        int imageWidth, imageHeight;
                        using (FileStream stream = new FileStream(graphFilePath, FileMode.Open))
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
                        Drawing drawing = InsertImage(mainPart.GetIdOfPart(imagePart), usedWidth, (int)(usedWidth * imageHeight / imageWidth));
                        body.AppendChild(new Paragraph(new Run(drawing)));
                    }
                    catch { }
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
        }
        private void addAgentTools()
        {
            AddHeading(content.Tools, "Heading1");

            var tools = content.GetResolvedToolInfos();
            if (tools.Count == 0)
            {
                body.AppendChild(new Paragraph(new Run(new Text("No tools configured."))));
                return;
            }

            // Summary table matching Copilot Studio UI columns
            Table summaryTable = CreateTable();
            summaryTable.Append(CreateHeaderRow(new Text("Name"), new Text("Type"), new Text("Available to"), new Text("Trigger"), new Text("Enabled")));
            foreach (AgentToolInfo tool in tools)
            {
                summaryTable.Append(CreateRow(
                    new Text(tool.Name),
                    new Text(tool.ToolType),
                    new Text(tool.AvailableTo ?? ""),
                    new Text(tool.Trigger ?? ""),
                    new Text(tool.Enabled ? "On" : "Off")));
            }
            body.Append(summaryTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Detail per tool
            foreach (AgentToolInfo tool in tools)
            {
                AddHeading("Tool: " + tool.Name, "Heading2");

                // Details table
                Table table = CreateTable();
                table.Append(CreateRow(new Text("Name"), new Text(tool.Name)));
                if (!string.IsNullOrEmpty(tool.Description))
                    table.Append(CreateRow(new Text("Description"), new Text(tool.Description)));
                table.Append(CreateRow(new Text("Type"), new Text(tool.ToolType)));
                table.Append(CreateRow(new Text("Available to"), new Text(tool.AvailableTo ?? "")));
                table.Append(CreateRow(new Text("Trigger"), new Text(tool.Trigger ?? "")));
                table.Append(CreateRow(new Text("Enabled"), new Text(tool.Enabled ? "On" : "Off")));
                if (!string.IsNullOrEmpty(tool.ConnectionReference))
                    table.Append(CreateRow(new Text("Connection Reference"), new Text(tool.ConnectionReference)));
                if (!string.IsNullOrEmpty(tool.OperationId))
                    table.Append(CreateRow(new Text("Operation"), new Text(tool.OperationId)));
                if (!string.IsNullOrEmpty(tool.FlowId))
                    table.Append(CreateRow(new Text("Flow ID"), new Text(tool.FlowId)));
                if (!string.IsNullOrEmpty(tool.AgentFlowName))
                    table.Append(CreateRow(new Text("Agent Flow"), new Text(tool.AgentFlowName)));
                if (!string.IsNullOrEmpty(tool.ModelParameters))
                    table.Append(CreateRow(new Text("Model Parameters"), new Text(tool.ModelParameters)));
                body.Append(table);
                body.AppendChild(new Paragraph(new Run(new Break())));

                // Inputs
                if (tool.Inputs.Count > 0)
                {
                    AddHeading("Inputs", "Heading3");

                    Table inputTable = CreateTable();
                    inputTable.Append(CreateHeaderRow(new Text("Input name"), new Text("Fill using"), new Text("Type"), new Text("Description")));
                    foreach (var input in tool.Inputs)
                    {
                        inputTable.Append(CreateRow(
                            new Text(input.Name + (input.IsRequired ? " *" : "")),
                            new Text(input.FillUsing ?? ""),
                            new Text(input.DataType ?? ""),
                            new Text(input.Description ?? "")));
                    }
                    body.Append(inputTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Outputs
                if (tool.Outputs.Count > 0)
                {
                    AddHeading("Outputs", "Heading3");

                    Table outputTable = CreateTable();
                    outputTable.Append(CreateHeaderRow(new Text("Output name"), new Text("Type"), new Text("Description")));
                    foreach (var output in tool.Outputs)
                    {
                        outputTable.Append(CreateRow(
                            new Text(output.Name),
                            new Text(output.DataType ?? ""),
                            new Text(output.Description ?? "")));
                    }
                    body.Append(outputTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Completion
                if (!string.IsNullOrEmpty(tool.ResponseActivity) || !string.IsNullOrEmpty(tool.OutputMode))
                {
                    AddHeading("Completion", "Heading3");

                    Table completionTable = CreateTable();
                    if (!string.IsNullOrEmpty(tool.ResponseActivity))
                        completionTable.Append(CreateRow(new Text("After running"), new Text("Send specific response")));
                    if (!string.IsNullOrEmpty(tool.ResponseMode))
                        completionTable.Append(CreateRow(new Text("Response Mode"), new Text(tool.ResponseMode)));
                    if (!string.IsNullOrEmpty(tool.OutputMode))
                        completionTable.Append(CreateRow(new Text("Output Mode"), new Text(tool.OutputMode)));
                    body.Append(completionTable);

                    if (!string.IsNullOrEmpty(tool.ResponseActivity))
                    {
                        AddHeading("Message to display:", "Heading4");
                        body.AppendChild(new Paragraph(new Run(new Text(tool.ResponseActivity))));
                    }
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Prompt text (for prompt tools)
                if (!string.IsNullOrEmpty(tool.PromptText))
                {
                    AddHeading("Prompt", "Heading3");
                    body.AppendChild(new Paragraph(new Run(new Text(tool.PromptText))));
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
        }

        private void addAgentEntities()
        {
            var entities = content.agent.GetEntities();
            if (entities.Count == 0) return;

            AddHeading(content.Entities, "Heading1");

            foreach (BotComponent entity in entities.OrderBy(e => e.Name))
            {
                AddHeading("Entity: " + entity.Name, "Heading2");

                string entityKind = entity.GetTopicKind();
                Table table = CreateTable();
                table.Append(CreateRow(new Text("Name"), new Text(entity.Name)));
                table.Append(CreateRow(new Text("Kind"), new Text(entityKind)));
                if (!string.IsNullOrEmpty(entity.Description))
                    table.Append(CreateRow(new Text("Description"), new Text(entity.Description)));

                if (entityKind == "RegexEntity")
                {
                    string pattern = entity.GetEntityPattern();
                    if (!string.IsNullOrEmpty(pattern))
                        table.Append(CreateRow(new Text("Pattern"), new Text(pattern)));
                }
                body.Append(table);
                body.AppendChild(new Paragraph(new Run(new Break())));

                if (entityKind == "ClosedListEntity")
                {
                    var items = entity.GetEntityItems();
                    if (items.Count > 0)
                    {
                        AddHeading("Items", "Heading3");

                        Table itemsTable = CreateTable();
                        itemsTable.Append(CreateHeaderRow(new Text("Display Name"), new Text("ID")));
                        foreach (var (id, displayName) in items)
                        {
                            itemsTable.Append(CreateRow(new Text(displayName), new Text(id)));
                        }
                        body.Append(itemsTable);
                        body.AppendChild(new Paragraph(new Run(new Break())));
                    }
                }
            }
        }

        private void addAgentVariables()
        {
            var variables = content.agent.GetVariables();
            if (variables.Count == 0) return;

            AddHeading(content.Variables, "Heading1");

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Name"), new Text("Scope"), new Text("Data Type"), new Text("AI Visibility"), new Text("External Init"), new Text("Description")));
            foreach (BotComponent variable in variables.OrderBy(v => v.Name))
            {
                var (scope, aiVisibility, dataType, isExternalInit) = variable.GetVariableDetails();
                table.Append(CreateRow(
                    new Text(variable.Name),
                    new Text(scope),
                    new Text(dataType),
                    new Text(aiVisibility),
                    new Text(isExternalInit ? "Yes" : "No"),
                    new Text(variable.Description ?? "")));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Global Variable Usage Tracking
            var usageMap = content.GetGlobalVariableUsageMap();
            if (usageMap.Count > 0)
            {
                AddHeading("Global Variable Usage", "Heading2");
                body.AppendChild(new Paragraph(new Run(new Text("Cross-reference showing which topics read or write each global variable."))));

                var varMetadata = new Dictionary<string, (string AIVisibility, string DataType)>(StringComparer.OrdinalIgnoreCase);
                foreach (BotComponent v in variables)
                {
                    var (_, aiVis, dt, _) = v.GetVariableDetails();
                    varMetadata["Global." + v.Name] = (aiVis, dt);
                }

                var sortedVars = usageMap.Keys.OrderByDescending(k =>
                {
                    if (varMetadata.TryGetValue(k, out var meta))
                        return meta.AIVisibility?.Contains("UseInAIContext", StringComparison.OrdinalIgnoreCase) == true
                            || meta.AIVisibility?.Contains("Public", StringComparison.OrdinalIgnoreCase) == true ? 1 : 0;
                    return 0;
                }).ThenBy(k => k);

                foreach (string varName in sortedVars)
                {
                    string header = varName;
                    if (varMetadata.TryGetValue(varName, out var meta)
                        && (meta.AIVisibility?.Contains("UseInAIContext", StringComparison.OrdinalIgnoreCase) == true
                            || meta.AIVisibility?.Contains("Public", StringComparison.OrdinalIgnoreCase) == true))
                    {
                        header += " (Orchestrator-visible)";
                    }
                    AddHeading(header, "Heading3");

                    Table usageTable = CreateTable();
                    usageTable.Append(CreateHeaderRow(new Text("Topic"), new Text("Access"), new Text("Context")));
                    foreach (var entry in usageMap[varName].OrderBy(e => e.AccessType).ThenBy(e => e.TopicName))
                    {
                        usageTable.Append(CreateRow(new Text(entry.TopicName), new Text(entry.AccessType.ToString()), new Text(entry.Context)));
                    }
                    body.Append(usageTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
        }

        private void addAgentChannels()
        {
            AddHeading("Channels", "Heading1");

            body.AppendChild(new Paragraph(new Run(new Text("Channels are not exported with the solution and are not documented automatically."))));
        }

        private static readonly string NotInExportMessage = "This setting is not available in the solution export.";

        private void addAgentSettings()
        {
            var config = content.agent.Configuration;
            var ai = config?.aISettings;

            AddHeading("Settings", "Heading1");

            // Generative AI
            AddHeading("Generative AI", "Heading2");

            Table genAiTable = CreateTable();
            genAiTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            genAiTable.Append(CreateRow(new Text("Generative Actions"), new Text(config?.settings?.GenerativeActionsEnabled == true ? "Enabled" : "Disabled")));
            genAiTable.Append(CreateRow(new Text("Use Model Knowledge"), new Text(ai?.useModelKnowledge == true ? "Yes" : "No")));
            genAiTable.Append(CreateRow(new Text("File Analysis"), new Text(ai?.isFileAnalysisEnabled == true ? "Enabled" : "Disabled")));
            genAiTable.Append(CreateRow(new Text("Semantic Search"), new Text(ai?.isSemanticSearchEnabled == true ? "Enabled" : "Disabled")));
            genAiTable.Append(CreateRow(new Text("Content Moderation"), new Text(ai?.contentModeration ?? "Unknown")));
            genAiTable.Append(CreateRow(new Text("Opt-in to Latest Models"), new Text(ai?.optInUseLatestModels == true ? "Yes" : "No")));
            body.Append(genAiTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Security
            AddHeading("Security", "Heading2");

            Table secTable = CreateTable();
            secTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            secTable.Append(CreateRow(new Text("Authentication Mode"), new Text(content.agent.GetAuthenticationModeDisplayName())));
            secTable.Append(CreateRow(new Text("Authentication Trigger"), new Text(content.agent.GetAuthenticationTriggerDisplayName())));
            body.Append(secTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Connection settings
            AddHeading("Connection settings", "Heading2");
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Authoring canvas
            AddHeading("Authoring canvas", "Heading2");
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Entities
            AddHeading("Entities", "Heading2");
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Skills
            AddHeading("Skills", "Heading2");
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Voice
            AddHeading("Voice", "Heading2");
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Languages
            AddHeading("Languages", "Heading2");

            Table langTable = CreateTable();
            langTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            langTable.Append(CreateRow(new Text("Primary Language"), new Text(content.agent.GetLanguageDisplayName())));
            body.Append(langTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Language understanding
            AddHeading("Language understanding", "Heading2");

            Table luTable = CreateTable();
            luTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            luTable.Append(CreateRow(new Text("Recognizer"), new Text(content.agent.GetRecognizerDisplayName())));
            body.Append(luTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Component collections
            AddHeading("Component collections", "Heading2");
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Advanced
            AddHeading("Advanced", "Heading2");

            Table advTable = CreateTable();
            advTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            advTable.Append(CreateRow(new Text("Template"), new Text(content.agent.Template ?? "")));
            advTable.Append(CreateRow(new Text("Runtime Provider"), new Text(content.agent.RuntimeProvider.ToString())));
            body.Append(advTable);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }
    }
}
