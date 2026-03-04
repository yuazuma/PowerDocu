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
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Agent - " + content.filename));
            ApplyStyleToParagraph("Heading1", para);

            Table table = CreateTable();
            table.Append(CreateRow(new Text("Agent Name"), new Text(content.agent.Name)));
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
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Details));
            ApplyStyleToParagraph("Heading2", para);

            // Description
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Description));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(CreateRunWithLinebreaks(content.agent.GetDescription() ?? "")));

            // Orchestration
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Orchestration));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(new Run(new Text($"{content.OrchestrationText} - {content.agent.GetOrchestration()}"))));

            // Response Model
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.ResponseModel));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(new Run(new Text(content.agent.GetResponseModel()))));

            // Instructions
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Instructions));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(CreateRunWithLinebreaks(content.agent.GetInstructions() ?? "")));

            // Knowledge
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Knowledge));
            ApplyStyleToParagraph("Heading3", para);
            var knowledgeSources = content.agent.GetKnowledge();
            var fileKnowledge = content.agent.GetFileKnowledge();
            if (knowledgeSources.Count > 0 || fileKnowledge.Count > 0)
            {
                Table knowledgeTable = CreateTable();
                knowledgeTable.Append(CreateHeaderRow(new Text("Name"), new Text("Source Type"), new Text("Details")));
                foreach (BotComponent ks in knowledgeSources)
                {
                    var (sourceKind, skillConfig) = ks.GetKnowledgeSourceDetails();
                    string site = ks.GetKnowledgeSourceSite();
                    string details = !string.IsNullOrEmpty(site) ? site : (!string.IsNullOrEmpty(skillConfig) ? skillConfig : "");
                    knowledgeTable.Append(CreateRow(new Text(ks.Name), new Text(sourceKind ?? ""), new Text(details)));
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
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Tools));
            ApplyStyleToParagraph("Heading3", para);
            var tools = content.agent.GetTools();
            if (tools.Count > 0)
            {
                foreach (BotComponent tool in tools.OrderBy(t => t.Name))
                {
                    body.AppendChild(new Paragraph(new Run(new Text(tool.Name))));
                }
            }
            else
            {
                body.AppendChild(new Paragraph(new Run(new Text("No tools configured."))));
            }

            // Triggers
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Triggers));
            ApplyStyleToParagraph("Heading3", para);
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
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Agents));
            ApplyStyleToParagraph("Heading3", para);
            body.AppendChild(new Paragraph(new Run(new Text("Sub-agents are not available in the solution export."))));

            // Topics
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Topics));
            ApplyStyleToParagraph("Heading3", para);
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name))
            {
                body.AppendChild(new Paragraph(new Run(new Text(topic.Name))));
            }

            // Suggested Prompts
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.SuggestedPrompts));
            ApplyStyleToParagraph("Heading3", para);
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

        private void addAgentTopicsOverview()
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Topics));
            ApplyStyleToParagraph("Heading1", para);

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
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Topic: " + topic.Name));
                ApplyStyleToParagraph("Heading2", para);

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
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Trigger Queries"));
                    ApplyStyleToParagraph("Heading3", para);

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
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Knowledge Source"));
                    ApplyStyleToParagraph("Heading3", para);

                    Table ksTable = CreateTable();
                    if (!string.IsNullOrEmpty(sourceKind))
                        ksTable.Append(CreateRow(new Text("Source Kind"), new Text(sourceKind)));
                    if (!string.IsNullOrEmpty(skillConfig))
                        ksTable.Append(CreateRow(new Text("Skill Configuration"), new Text(skillConfig)));
                    body.Append(ksTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Variables
                var variables = topic.GetTopicVariables();
                if (variables.Count > 0)
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Variables"));
                    ApplyStyleToParagraph("Heading3", para);

                    Table varTable = CreateTable();
                    varTable.Append(CreateHeaderRow(new Text("Variable"), new Text("Context")));
                    foreach (var (variable, context) in variables)
                    {
                        varTable.Append(CreateRow(new Text(variable), new Text(context)));
                    }
                    body.Append(varTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                // Topic flow diagram
                string graphFile = topic.getTopicFileName() + "-detailed.png";
                string graphFilePath = Path.Combine(content.folderPath, "Topics", graphFile);
                if (File.Exists(graphFilePath))
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Topic Flow"));
                    ApplyStyleToParagraph("Heading3", para);

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
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Tools));
            ApplyStyleToParagraph("Heading1", para);

            var tools = content.agent.GetTools();
            if (tools.Count == 0)
            {
                body.AppendChild(new Paragraph(new Run(new Text("No tools configured."))));
                return;
            }

            foreach (BotComponent tool in tools.OrderBy(t => t.Name))
            {
                para = body.AppendChild(new Paragraph());
                run = para.AppendChild(new Run());
                run.AppendChild(new Text("Tool: " + tool.Name));
                ApplyStyleToParagraph("Heading2", para);

                var (actionKind, connectionRef, operationId, flowId, modelDisplayName, inputs, outputs) = tool.GetToolDetails();

                Table table = CreateTable();
                table.Append(CreateRow(new Text("Name"), new Text(tool.Name)));
                if (!string.IsNullOrEmpty(modelDisplayName))
                    table.Append(CreateRow(new Text("Display Name"), new Text(modelDisplayName)));
                if (!string.IsNullOrEmpty(tool.Description))
                    table.Append(CreateRow(new Text("Description"), new Text(tool.Description)));

                string actionTypeDisplay = actionKind switch
                {
                    "InvokeConnectorTaskAction" => "Connector",
                    "InvokeFlowTaskAction" => "Power Automate Flow",
                    _ => actionKind
                };
                table.Append(CreateRow(new Text("Action Type"), new Text(actionTypeDisplay)));

                if (!string.IsNullOrEmpty(connectionRef))
                    table.Append(CreateRow(new Text("Connection Reference"), new Text(connectionRef)));
                if (!string.IsNullOrEmpty(operationId))
                    table.Append(CreateRow(new Text("Operation"), new Text(operationId)));
                if (!string.IsNullOrEmpty(flowId))
                    table.Append(CreateRow(new Text("Flow ID"), new Text(flowId)));
                body.Append(table);
                body.AppendChild(new Paragraph(new Run(new Break())));

                if (inputs.Count > 0)
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Inputs"));
                    ApplyStyleToParagraph("Heading3", para);

                    Table inputTable = CreateTable();
                    inputTable.Append(CreateHeaderRow(new Text("Input")));
                    foreach (string input in inputs)
                    {
                        inputTable.Append(CreateRow(new Text(input)));
                    }
                    body.Append(inputTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }

                if (outputs.Count > 0)
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Outputs"));
                    ApplyStyleToParagraph("Heading3", para);

                    Table outputTable = CreateTable();
                    outputTable.Append(CreateHeaderRow(new Text("Output")));
                    foreach (string output in outputs)
                    {
                        outputTable.Append(CreateRow(new Text(output)));
                    }
                    body.Append(outputTable);
                    body.AppendChild(new Paragraph(new Run(new Break())));
                }
            }
        }

        private void addAgentEntities()
        {
            var entities = content.agent.GetEntities();
            if (entities.Count == 0) return;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Entities));
            ApplyStyleToParagraph("Heading1", para);

            foreach (BotComponent entity in entities.OrderBy(e => e.Name))
            {
                para = body.AppendChild(new Paragraph());
                run = para.AppendChild(new Run());
                run.AppendChild(new Text("Entity: " + entity.Name));
                ApplyStyleToParagraph("Heading2", para);

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
                        para = body.AppendChild(new Paragraph());
                        run = para.AppendChild(new Run());
                        run.AppendChild(new Text("Items"));
                        ApplyStyleToParagraph("Heading3", para);

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

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.Variables));
            ApplyStyleToParagraph("Heading1", para);

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Name"), new Text("Scope"), new Text("Data Type"), new Text("AI Visibility"), new Text("External Init")));
            foreach (BotComponent variable in variables.OrderBy(v => v.Name))
            {
                var (scope, aiVisibility, dataType, isExternalInit) = variable.GetVariableDetails();
                table.Append(CreateRow(
                    new Text(variable.Name),
                    new Text(scope),
                    new Text(dataType),
                    new Text(aiVisibility),
                    new Text(isExternalInit ? "Yes" : "No")));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAgentChannels()
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Channels"));
            ApplyStyleToParagraph("Heading1", para);

            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Channels are not exported with the solution and are not documented automatically."));
        }

        private static readonly string NotInExportMessage = "This setting is not available in the solution export.";

        private void addAgentSettings()
        {
            var config = content.agent.Configuration;
            var ai = config?.aISettings;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Settings"));
            ApplyStyleToParagraph("Heading1", para);

            // Generative AI
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Generative AI"));
            ApplyStyleToParagraph("Heading2", para);

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
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Security"));
            ApplyStyleToParagraph("Heading2", para);

            Table secTable = CreateTable();
            secTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            secTable.Append(CreateRow(new Text("Authentication Mode"), new Text(content.agent.GetAuthenticationModeDisplayName())));
            secTable.Append(CreateRow(new Text("Authentication Trigger"), new Text(content.agent.GetAuthenticationTriggerDisplayName())));
            body.Append(secTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Connection settings
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Connection settings"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Authoring canvas
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Authoring canvas"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Entities
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Entities"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Skills
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Skills"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Voice
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Voice"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Languages
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Languages"));
            ApplyStyleToParagraph("Heading2", para);

            Table langTable = CreateTable();
            langTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            langTable.Append(CreateRow(new Text("Primary Language"), new Text(content.agent.GetLanguageDisplayName())));
            body.Append(langTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Language understanding
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Language understanding"));
            ApplyStyleToParagraph("Heading2", para);

            Table luTable = CreateTable();
            luTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            luTable.Append(CreateRow(new Text("Recognizer"), new Text(content.agent.GetRecognizerDisplayName())));
            body.Append(luTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Component collections
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Component collections"));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(new Text(NotInExportMessage))));

            // Advanced
            para = body.AppendChild(new Paragraph());
            run = para.AppendChild(new Run());
            run.AppendChild(new Text("Advanced"));
            ApplyStyleToParagraph("Heading2", para);

            Table advTable = CreateTable();
            advTable.Append(CreateHeaderRow(new Text("Setting"), new Text("Value")));
            advTable.Append(CreateRow(new Text("Template"), new Text(content.agent.Template ?? "")));
            advTable.Append(CreateRow(new Text("Runtime Provider"), new Text(content.agent.RuntimeProvider.ToString())));
            body.Append(advTable);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }
    }
}
