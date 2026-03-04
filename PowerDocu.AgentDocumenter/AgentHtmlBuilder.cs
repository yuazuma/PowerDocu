using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using PowerDocu.Common;

namespace PowerDocu.AgentDocumenter
{
    class AgentHtmlBuilder : HtmlBuilder
    {
        private readonly AgentDocumentationContent content;
        private readonly string mainFileName, knowledgeFileName, toolsFileName, agentsFileName, topicsFileName, channelsFileName, settingsFileName, entitiesFileName, variablesFileName;
        private readonly Dictionary<string, string> topicFileNames = new Dictionary<string, string>();

        public AgentHtmlBuilder(AgentDocumentationContent contentdocumentation)
        {
            content = contentdocumentation;
            Directory.CreateDirectory(content.folderPath);
            WriteDefaultStylesheet(content.folderPath);

            mainFileName = ("index-" + content.filename + ".html").Replace(" ", "-");
            knowledgeFileName = ("knowledge-" + content.filename + ".html").Replace(" ", "-");
            toolsFileName = ("tools-" + content.filename + ".html").Replace(" ", "-");
            agentsFileName = ("agents-" + content.filename + ".html").Replace(" ", "-");
            topicsFileName = ("topics-" + content.filename + ".html").Replace(" ", "-");
            channelsFileName = ("channels-" + content.filename + ".html").Replace(" ", "-");
            settingsFileName = ("settings-" + content.filename + ".html").Replace(" ", "-");
            entitiesFileName = ("entities-" + content.filename + ".html").Replace(" ", "-");
            variablesFileName = ("variables-" + content.filename + ".html").Replace(" ", "-");

            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                topicFileNames[topic.Name] = ("topic-" + CharsetHelper.GetSafeName(topic.Name) + "-" + content.filename + ".html").Replace(" ", "-");
            }

            addAgentOverview();
            addAgentKnowledgeInfo();
            addAgentTools();
            addAgentAgentsInfo();
            addAgentTopics();
            addAgentEntities();
            addAgentVariables();
            addAgentChannels();
            addAgentSettings();
            NotificationHelper.SendNotification("Created HTML documentation for " + content.filename);
        }

        private string getNavigationHtml(bool isSubfolder = false)
        {
            string prefix = isSubfolder ? "../" : "";
            var navItems = new List<(string label, string href)>
            {
                ("Overview", prefix + mainFileName),
                ("Knowledge", prefix + knowledgeFileName),
                ("Tools", prefix + toolsFileName),
                ("Entities", prefix + entitiesFileName),
                ("Variables", prefix + variablesFileName),
                ("Agents", prefix + agentsFileName),
                ("Topics", prefix + topicsFileName),
                ("Channels", prefix + channelsFileName),
                ("Settings", prefix + settingsFileName)
            };
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"<div class=\"nav-title\">{Encode(content.filename)}</div>");
            sb.Append(NavigationList(navItems));
            return sb.ToString();
        }

        private string buildMetadataTable(bool isSubfolder = false)
        {
            string prefix = isSubfolder ? "../" : "";
            StringBuilder sb = new StringBuilder();
            sb.Append(TableStart("Property", "Value"));
            sb.Append(TableRow("Agent Name", content.agent.Name));
            if (!String.IsNullOrEmpty(content.agent.IconBase64))
            {
                Directory.CreateDirectory(content.folderPath);
                Bitmap agentLogo = ImageHelper.ConvertBase64ToBitmap(content.agent.IconBase64);
                string logoFileName = $"agentlogo-{content.filename.Replace(" ", "-")}.png";
                agentLogo.Save(content.folderPath + logoFileName);
                sb.Append(TableRowRaw("Agent Logo", Image("Agent Logo", prefix + logoFileName)));
                agentLogo.Dispose();
            }
            sb.Append(TableRow(content.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion()));
            sb.AppendLine(TableEnd());
            return sb.ToString();
        }

        private void addAgentOverview()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());

            body.AppendLine(Heading(2, content.Details));
            body.AppendLine(Heading(3, content.Description));
            body.AppendLine(ParagraphWithLinebreaks(content.agent.GetDescription()));
            body.AppendLine(Heading(3, content.Orchestration));
            body.AppendLine(Paragraph($"{content.OrchestrationText} - {content.agent.GetOrchestration()}"));
            body.AppendLine(Heading(3, content.ResponseModel));
            body.AppendLine(Paragraph(content.agent.GetResponseModel()));
            body.AppendLine(Heading(3, content.Instructions));
            body.AppendLine(ParagraphWithLinebreaks(content.agent.GetInstructions()));

            body.AppendLine(Heading(3, content.Knowledge));
            var knowledgeSources = content.agent.GetKnowledge();
            var fileKnowledge = content.agent.GetFileKnowledge();
            if (knowledgeSources.Count > 0 || fileKnowledge.Count > 0)
            {
                body.Append(TableStart("Name", "Source Type", "Details"));
                foreach (BotComponent ks in knowledgeSources)
                {
                    var (sourceKind, skillConfig) = ks.GetKnowledgeSourceDetails();
                    string site = ks.GetKnowledgeSourceSite();
                    string details = !string.IsNullOrEmpty(site) ? site : (!string.IsNullOrEmpty(skillConfig) ? skillConfig : "");
                    body.Append(TableRow(ks.Name, sourceKind ?? "", details));
                }
                foreach (BotComponent fk in fileKnowledge)
                {
                    string mimeType = !string.IsNullOrEmpty(fk.FileDataMimeType) ? $" ({fk.FileDataMimeType})" : "";
                    body.Append(TableRow(fk.Name, "File" + mimeType, fk.FileDataName ?? ""));
                }
                body.AppendLine(TableEnd());
            }
            else
            {
                body.AppendLine(Paragraph("No knowledge sources configured."));
            }

            body.AppendLine(Heading(3, content.Tools));
            var overviewTools = content.agent.GetTools();
            if (overviewTools.Count > 0)
            {
                body.AppendLine(BulletListStart());
                foreach (BotComponent tool in overviewTools.OrderBy(t => t.Name))
                {
                    body.AppendLine(BulletItem(tool.Name));
                }
                body.AppendLine(BulletListEnd());
            }
            else
            {
                body.AppendLine(Paragraph("No tools configured."));
            }

            body.AppendLine(Heading(3, content.Triggers));
            var overviewTriggers = content.agent.GetTriggers();
            if (overviewTriggers.Count > 0)
            {
                body.AppendLine(BulletListStart());
                foreach (BotComponent trigger in overviewTriggers.OrderBy(t => t.Name))
                {
                    var (triggerKind, flowId, connectionType) = trigger.GetTriggerDetails();
                    string triggerInfo = !string.IsNullOrEmpty(connectionType) ? $" ({connectionType})" : "";
                    body.AppendLine(BulletItem(trigger.Name + triggerInfo));
                }
                body.AppendLine(BulletListEnd());
            }
            else
            {
                body.AppendLine(Paragraph("No triggers configured."));
            }

            body.AppendLine(Heading(3, content.Agents));
            body.AppendLine(Paragraph("Sub-agents are not available in the solution export."));

            body.AppendLine(Heading(3, content.Topics));
            body.AppendLine(BulletListStart());
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name))
            {
                string topicFile = topicFileNames.GetValueOrDefault(topic.Name, "#");
                body.AppendLine(BulletItemRaw(Link(topic.Name, "Topics/" + topicFile)));
            }
            body.AppendLine(BulletListEnd());

            body.AppendLine(Heading(3, content.SuggestedPrompts));
            body.AppendLine(Paragraph(content.SuggestedPromptsText));
            body.Append(TableStart("Prompt Title", "Prompt"));
            Dictionary<string, string> conversationStarters = content.agent.GetSuggestedPrompts();
            foreach (var kvp in conversationStarters.OrderBy(x => x.Key))
            {
                body.Append(TableRow(kvp.Key, kvp.Value));
            }
            body.AppendLine(TableEnd());

            SaveHtmlFile(Path.Combine(content.folderPath, mainFileName),
                WrapInHtmlPage($"Agent - {content.filename}", body.ToString(), getNavigationHtml()));
        }

        private void addAgentKnowledgeInfo()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.Knowledge));
            body.AppendLine(Paragraph("Knowledge sources for this agent."));

            var knowledgeSources = content.agent.GetKnowledge();
            var fileKnowledge = content.agent.GetFileKnowledge();
            if (knowledgeSources.Count > 0 || fileKnowledge.Count > 0)
            {
                body.Append(TableStart("Name", "Source Type", "Details", "Description"));
                foreach (BotComponent ks in knowledgeSources)
                {
                    var (sourceKind, skillConfig) = ks.GetKnowledgeSourceDetails();
                    string site = ks.GetKnowledgeSourceSite();
                    string details = !string.IsNullOrEmpty(site) ? site : (!string.IsNullOrEmpty(skillConfig) ? skillConfig : "");
                    body.Append(TableRow(ks.Name, sourceKind ?? "", details, ks.Description ?? ""));
                }
                foreach (BotComponent fk in fileKnowledge)
                {
                    string mimeType = !string.IsNullOrEmpty(fk.FileDataMimeType) ? $" ({fk.FileDataMimeType})" : "";
                    body.Append(TableRow(fk.Name, "File" + mimeType, fk.FileDataName ?? "", fk.Description ?? ""));
                }
                body.AppendLine(TableEnd());
            }
            else
            {
                body.AppendLine(Paragraph("No knowledge sources configured."));
            }

            SaveHtmlFile(Path.Combine(content.folderPath, knowledgeFileName),
                WrapInHtmlPage($"Knowledge - {content.filename}", body.ToString(), getNavigationHtml()));
        }

        private void addAgentTools()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.Tools));

            var tools = content.agent.GetTools();
            if (tools.Count == 0)
            {
                body.AppendLine(Paragraph("No tools configured."));
            }
            else
            {
                // Summary table
                body.Append(TableStart("Name", "Action Type", "Connection", "Operation"));
                foreach (BotComponent tool in tools.OrderBy(t => t.Name))
                {
                    var (actionKind, connectionRef, operationId, flowId, modelDisplayName, inputs, outputs) = tool.GetToolDetails();
                    string actionTypeDisplay = actionKind switch
                    {
                        "InvokeConnectorTaskAction" => "Connector",
                        "InvokeFlowTaskAction" => "Power Automate Flow",
                        _ => actionKind
                    };
                    body.Append(TableRow(tool.Name, actionTypeDisplay, connectionRef ?? "", operationId ?? ""));
                }
                body.AppendLine(TableEnd());

                // Detail per tool
                foreach (BotComponent tool in tools.OrderBy(t => t.Name))
                {
                    body.AppendLine(Heading(3, tool.Name));
                    var (actionKind, connectionRef, operationId, flowId, modelDisplayName, inputs, outputs) = tool.GetToolDetails();
                    body.Append(TableStart("Property", "Value"));
                    if (!string.IsNullOrEmpty(modelDisplayName))
                        body.Append(TableRow("Display Name", modelDisplayName));
                    if (!string.IsNullOrEmpty(tool.Description))
                        body.Append(TableRow("Description", tool.Description));
                    string actionTypeDisplay = actionKind switch
                    {
                        "InvokeConnectorTaskAction" => "Connector",
                        "InvokeFlowTaskAction" => "Power Automate Flow",
                        _ => actionKind
                    };
                    body.Append(TableRow("Action Type", actionTypeDisplay));
                    if (!string.IsNullOrEmpty(connectionRef))
                        body.Append(TableRow("Connection Reference", connectionRef));
                    if (!string.IsNullOrEmpty(operationId))
                        body.Append(TableRow("Operation", operationId));
                    if (!string.IsNullOrEmpty(flowId))
                        body.Append(TableRow("Flow ID", flowId));
                    body.AppendLine(TableEnd());

                    if (inputs.Count > 0)
                    {
                        body.AppendLine(Heading(4, "Inputs"));
                        body.Append(TableStart("Input"));
                        foreach (string input in inputs)
                        {
                            body.Append(TableRow(input));
                        }
                        body.AppendLine(TableEnd());
                    }

                    if (outputs.Count > 0)
                    {
                        body.AppendLine(Heading(4, "Outputs"));
                        body.Append(TableStart("Output"));
                        foreach (string output in outputs)
                        {
                            body.Append(TableRow(output));
                        }
                        body.AppendLine(TableEnd());
                    }
                }
            }

            SaveHtmlFile(Path.Combine(content.folderPath, toolsFileName),
                WrapInHtmlPage($"Tools - {content.filename}", body.ToString(), getNavigationHtml()));
        }

        private void addAgentEntities()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.Entities));

            var entities = content.agent.GetEntities();
            if (entities.Count == 0)
            {
                body.AppendLine(Paragraph("No custom entities defined."));
            }
            else
            {
                foreach (BotComponent entity in entities.OrderBy(e => e.Name))
                {
                    string entityKind = entity.GetTopicKind();
                    body.AppendLine(Heading(3, entity.Name));

                    body.Append(TableStart("Property", "Value"));
                    body.Append(TableRow("Name", entity.Name));
                    body.Append(TableRow("Kind", entityKind));
                    if (!string.IsNullOrEmpty(entity.Description))
                        body.Append(TableRow("Description", entity.Description));
                    if (entityKind == "RegexEntity")
                    {
                        string pattern = entity.GetEntityPattern();
                        if (!string.IsNullOrEmpty(pattern))
                            body.Append(TableRow("Pattern", pattern));
                    }
                    body.AppendLine(TableEnd());

                    if (entityKind == "ClosedListEntity")
                    {
                        var items = entity.GetEntityItems();
                        if (items.Count > 0)
                        {
                            body.AppendLine(Heading(4, "Items"));
                            body.Append(TableStart("Display Name", "ID"));
                            foreach (var (id, displayName) in items)
                            {
                                body.Append(TableRow(displayName, id));
                            }
                            body.AppendLine(TableEnd());
                        }
                    }
                }
            }

            SaveHtmlFile(Path.Combine(content.folderPath, entitiesFileName),
                WrapInHtmlPage($"Entities - {content.filename}", body.ToString(), getNavigationHtml()));
        }

        private void addAgentVariables()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.Variables));

            var variables = content.agent.GetVariables();
            if (variables.Count == 0)
            {
                body.AppendLine(Paragraph("No global variables defined."));
            }
            else
            {
                body.Append(TableStart("Name", "Scope", "Data Type", "AI Visibility", "External Init", "Description"));
                foreach (BotComponent variable in variables.OrderBy(v => v.Name))
                {
                    var (scope, aiVisibility, dataType, isExternalInit) = variable.GetVariableDetails();
                    body.Append(TableRow(variable.Name, scope, dataType, aiVisibility, isExternalInit ? "Yes" : "No", variable.Description ?? ""));
                }
                body.AppendLine(TableEnd());
            }

            SaveHtmlFile(Path.Combine(content.folderPath, variablesFileName),
                WrapInHtmlPage($"Variables - {content.filename}", body.ToString(), getNavigationHtml()));
        }

        private void addAgentAgentsInfo()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.Agents));
            body.AppendLine(Paragraph("Sub-agents are not available in the solution export."));
            SaveHtmlFile(Path.Combine(content.folderPath, agentsFileName),
                WrapInHtmlPage($"Agents - {content.filename}", body.ToString(), getNavigationHtml()));
        }

        private void addAgentTopics()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.Topics));
            body.Append(TableStart("Name", "Type", "Trigger", "Kind"));
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                string topicFile = topicFileNames.GetValueOrDefault(topic.Name, "#");
                string topicType = topic.GetComponentTypeDisplayName();
                string triggerType = topic.GetTriggerTypeForTopic();
                string topicKind = topic.GetTopicKind() == "KnowledgeSourceConfiguration" ? "Knowledge" : topic.GetTopicKind();
                body.Append(TableRowRaw(Link(topic.Name, "Topics/" + topicFile), Encode(topicType), Encode(triggerType), Encode(topicKind)));
            }
            body.AppendLine(TableEnd());

            SaveHtmlFile(Path.Combine(content.folderPath, topicsFileName),
                WrapInHtmlPage($"Topics - {content.filename}", body.ToString(), getNavigationHtml()));

            // Build per-topic pages
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                string topicFile = topicFileNames.GetValueOrDefault(topic.Name, topic.Name + ".html");
                buildTopicPage(topic, topicFile);
            }
        }

        private void buildTopicPage(BotComponent topic, string topicFile)
        {
            StringBuilder topicBody = new StringBuilder();
            topicBody.AppendLine(Heading(1, $"Agent - {content.filename}"));
            topicBody.AppendLine(buildMetadataTable(true));
            topicBody.AppendLine(Heading(2, "Topic: " + topic.Name));

            // Metadata table
            topicBody.Append(TableStart("Property", "Value"));
            topicBody.Append(TableRow("Name", topic.Name));
            topicBody.Append(TableRow("Type", topic.GetComponentTypeDisplayName()));
            topicBody.Append(TableRow("Trigger", topic.GetTriggerTypeForTopic()));
            topicBody.Append(TableRow("Topic Kind", topic.GetTopicKind()));
            if (!string.IsNullOrEmpty(topic.Description))
            {
                topicBody.Append(TableRow("Description", topic.Description));
            }
            string modelDesc = topic.GetModelDescription();
            if (!string.IsNullOrEmpty(modelDesc))
            {
                topicBody.Append(TableRow("Model Description", modelDesc));
            }
            string startBehavior = topic.GetStartBehavior();
            if (!string.IsNullOrEmpty(startBehavior))
            {
                topicBody.Append(TableRow("Start Behavior", startBehavior));
            }
            topicBody.AppendLine(TableEnd());

            // Trigger queries
            List<string> triggerQueries = topic.GetTriggerQueries();
            if (triggerQueries.Count > 0)
            {
                topicBody.AppendLine(Heading(3, "Trigger Queries"));
                topicBody.AppendLine(BulletListStart());
                foreach (string query in triggerQueries)
                {
                    topicBody.AppendLine(BulletItem(query));
                }
                topicBody.AppendLine(BulletListEnd());
            }

            // Knowledge source details
            if (topic.GetTopicKind() == "KnowledgeSourceConfiguration")
            {
                var (sourceKind, skillConfig) = topic.GetKnowledgeSourceDetails();
                topicBody.AppendLine(Heading(3, "Knowledge Source"));
                topicBody.Append(TableStart("Property", "Value"));
                if (!string.IsNullOrEmpty(sourceKind))
                    topicBody.Append(TableRow("Source Kind", sourceKind));
                if (!string.IsNullOrEmpty(skillConfig))
                    topicBody.Append(TableRow("Skill Configuration", skillConfig));
                topicBody.AppendLine(TableEnd());
            }

            // Variables
            var variables = topic.GetTopicVariables();
            if (variables.Count > 0)
            {
                topicBody.AppendLine(Heading(3, "Variables"));
                topicBody.Append(TableStart("Variable", "Context"));
                foreach (var (variable, context) in variables)
                {
                    topicBody.Append(TableRow(variable, context));
                }
                topicBody.AppendLine(TableEnd());
            }

            // Topic flow diagram
            string graphFile = topic.getTopicFileName() + "-detailed.svg";
            if (File.Exists(Path.Combine(content.folderPath, "Topics", graphFile)))
            {
                topicBody.AppendLine(Heading(3, "Topic Flow"));
                topicBody.AppendLine(ParagraphRaw(Image("Topic Flow Diagram", graphFile)));
            }

            Directory.CreateDirectory(Path.Combine(content.folderPath, "Topics"));
            SaveHtmlFile(Path.Combine(content.folderPath, "Topics", topicFile),
                WrapInHtmlPage($"Topic: {topic.Name}", topicBody.ToString(), getNavigationHtml(true), "../style.css"));
        }

        private void addAgentChannels()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, "Channels"));
            body.AppendLine(Paragraph("Channels are not exported with the solution and are not documented automatically."));
            SaveHtmlFile(Path.Combine(content.folderPath, channelsFileName),
                WrapInHtmlPage($"Channels - {content.filename}", body.ToString(), getNavigationHtml()));
        }

        private static readonly string NotInExportMessage = "This setting is not available in the solution export.";

        private void addAgentSettings()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());

            var config = content.agent.Configuration;
            var ai = config?.aISettings;

            // Generative AI
            body.AppendLine(Heading(2, "Generative AI"));
            body.Append(TableStart("Setting", "Value"));
            body.Append(TableRow("Generative Actions", config?.settings?.GenerativeActionsEnabled == true ? "Enabled" : "Disabled"));
            body.Append(TableRow("Use Model Knowledge", ai?.useModelKnowledge == true ? "Yes" : "No"));
            body.Append(TableRow("File Analysis", ai?.isFileAnalysisEnabled == true ? "Enabled" : "Disabled"));
            body.Append(TableRow("Semantic Search", ai?.isSemanticSearchEnabled == true ? "Enabled" : "Disabled"));
            body.Append(TableRow("Content Moderation", ai?.contentModeration ?? "Unknown"));
            body.Append(TableRow("Opt-in to Latest Models", ai?.optInUseLatestModels == true ? "Yes" : "No"));
            body.AppendLine(TableEnd());

            // Security
            body.AppendLine(Heading(2, "Security"));
            body.Append(TableStart("Setting", "Value"));
            body.Append(TableRow("Authentication Mode", content.agent.GetAuthenticationModeDisplayName()));
            body.Append(TableRow("Authentication Trigger", content.agent.GetAuthenticationTriggerDisplayName()));
            body.AppendLine(TableEnd());

            // Connection settings
            body.AppendLine(Heading(2, "Connection settings"));
            body.AppendLine(Paragraph(NotInExportMessage));

            // Authoring canvas
            body.AppendLine(Heading(2, "Authoring canvas"));
            body.AppendLine(Paragraph(NotInExportMessage));

            // Entities
            body.AppendLine(Heading(2, "Entities"));
            body.AppendLine(Paragraph(NotInExportMessage));

            // Skills
            body.AppendLine(Heading(2, "Skills"));
            body.AppendLine(Paragraph(NotInExportMessage));

            // Voice
            body.AppendLine(Heading(2, "Voice"));
            body.AppendLine(Paragraph(NotInExportMessage));

            // Languages
            body.AppendLine(Heading(2, "Languages"));
            body.Append(TableStart("Setting", "Value"));
            body.Append(TableRow("Primary Language", content.agent.GetLanguageDisplayName()));
            body.AppendLine(TableEnd());

            // Language understanding
            body.AppendLine(Heading(2, "Language understanding"));
            body.Append(TableStart("Setting", "Value"));
            body.Append(TableRow("Recognizer", content.agent.GetRecognizerDisplayName()));
            body.AppendLine(TableEnd());

            // Component collections
            body.AppendLine(Heading(2, "Component collections"));
            body.AppendLine(Paragraph(NotInExportMessage));

            // Advanced
            body.AppendLine(Heading(2, "Advanced"));
            body.Append(TableStart("Setting", "Value"));
            body.Append(TableRow("Template", content.agent.Template ?? ""));
            body.Append(TableRow("Runtime Provider", content.agent.RuntimeProvider.ToString()));
            body.AppendLine(TableEnd());

            SaveHtmlFile(Path.Combine(content.folderPath, settingsFileName),
                WrapInHtmlPage($"Settings - {content.filename}", body.ToString(), getNavigationHtml()));
        }
    }
}
