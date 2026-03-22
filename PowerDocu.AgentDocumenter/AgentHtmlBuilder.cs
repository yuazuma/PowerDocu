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
        private readonly Dictionary<string, string> knowledgeDetailFileNames = new Dictionary<string, string>();

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
            foreach (BotComponent ks in content.agent.GetKnowledge().Concat(content.agent.GetFileKnowledge()).OrderBy(k => k.Name))
            {
                knowledgeDetailFileNames[ks.Name] = ("knowledge-" + CharsetHelper.GetSafeName(ks.Name) + "-" + content.filename + ".html").Replace(" ", "-");
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
            var navItems = new List<(string label, string href)>();
            if (content.context?.Solution != null)
            {
                string solutionPrefix = isSubfolder ? "../../" : "../";
                if (content.context?.Config?.documentSolution == true)
                    navItems.Add(("Solution", solutionPrefix + CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName)));
                else
                    navItems.Add((content.context.Solution.UniqueName, ""));
            }
            navItems.AddRange(new (string label, string href)[]
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
            });
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
            string responseModelText = content.agent.GetResponseModelDisplayName();
            string modelHint = content.agent.GetModelNameHint();
            if (!string.IsNullOrEmpty(modelHint))
                responseModelText += $" (Model: {modelHint})";
            body.AppendLine(Paragraph(responseModelText));
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
                    string detailFile = knowledgeDetailFileNames.GetValueOrDefault(ks.Name, "#");
                    string details = content.agent.GetKnowledgeDetailsSummary(ks);
                    string site = ks.GetKnowledgeSourceSite();
                    string detailsCell = !string.IsNullOrEmpty(site) ? Link(details, site) : Encode(details);
                    body.Append(TableRowRaw(Link(ks.Name, "Knowledge/" + detailFile), Encode(ks.GetSourceKindDisplayName()), detailsCell));
                }
                foreach (BotComponent fk in fileKnowledge)
                {
                    string detailFile = knowledgeDetailFileNames.GetValueOrDefault(fk.Name, "#");
                    string mimeType = !string.IsNullOrEmpty(fk.FileDataMimeType) ? $" ({fk.FileDataMimeType})" : "";
                    body.Append(TableRowRaw(Link(fk.Name, "Knowledge/" + detailFile), Encode("File" + mimeType), Encode(fk.FileDataName ?? "")));
                }
                body.AppendLine(TableEnd());
            }
            else
            {
                body.AppendLine(Paragraph("No knowledge sources configured."));
            }

            body.AppendLine(Heading(3, content.Tools));
            var overviewTools = content.GetResolvedToolInfos();
            if (overviewTools.Count > 0)
            {
                body.AppendLine(BulletListStart());
                foreach (AgentToolInfo tool in overviewTools)
                {
                    string anchor = SanitizeAnchorId(tool.Name);
                    body.AppendLine(BulletItemRaw(Link($"{tool.Name} ({tool.ToolType})", toolsFileName + "#" + anchor)));
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
            var overviewAgents = content.GetResolvedConnectedAgentInfos();
            if (overviewAgents.Count > 0)
            {
                body.AppendLine(BulletListStart());
                foreach (var agentInfo in overviewAgents.OrderBy(a => a.Name))
                {
                    string suffix = " (" + agentInfo.ConnectionType + ")";
                    string safeName = CharsetHelper.GetSafeName(agentInfo.Name);
                    string agentFolder = "AgentDoc " + safeName;
                    string agentFile = ("index-" + safeName + ".html").Replace(" ", "-");
                    string agentLink = "../" + agentFolder + "/" + agentFile;
                    body.AppendLine(BulletItemRaw(Link(agentInfo.Name, agentLink) + Encode(suffix)));
                }
                body.AppendLine(BulletListEnd());
            }
            else
            {
                body.AppendLine(Paragraph("No connected agents configured."));
            }

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
                body.Append(TableStart("Name", "Source Type", "Official Source", "Details", "Description"));
                foreach (BotComponent ks in knowledgeSources)
                {
                    string detailFile = knowledgeDetailFileNames.GetValueOrDefault(ks.Name, "#");
                    string details = content.agent.GetKnowledgeDetailsSummary(ks);
                    string officialSource = ks.GetOfficialSourceDisplayName();
                    string descriptionPreview = !string.IsNullOrEmpty(ks.Description) && ks.Description.Length > 100
                        ? ks.Description.Substring(0, 100) + "..."
                        : ks.Description ?? "";
                    string site = ks.GetKnowledgeSourceSite();
                    string detailsCell = !string.IsNullOrEmpty(site) ? Link(details, site) : Encode(details);
                    body.Append(TableRowRaw(
                        Link(ks.Name, "Knowledge/" + detailFile),
                        Encode(ks.GetSourceKindDisplayName()),
                        Encode(officialSource),
                        detailsCell,
                        Encode(descriptionPreview)));
                }
                foreach (BotComponent fk in fileKnowledge)
                {
                    string detailFile = knowledgeDetailFileNames.GetValueOrDefault(fk.Name, "#");
                    string mimeType = !string.IsNullOrEmpty(fk.FileDataMimeType) ? $" ({fk.FileDataMimeType})" : "";
                    string descriptionPreview = !string.IsNullOrEmpty(fk.Description) && fk.Description.Length > 100
                        ? fk.Description.Substring(0, 100) + "..."
                        : fk.Description ?? "";
                    body.Append(TableRowRaw(
                        Link(fk.Name, "Knowledge/" + detailFile),
                        Encode("File" + mimeType),
                        Encode(""),
                        Encode(fk.FileDataName ?? ""),
                        Encode(descriptionPreview)));
                }
                body.AppendLine(TableEnd());
            }
            else
            {
                body.AppendLine(Paragraph("No knowledge sources configured."));
            }

            SaveHtmlFile(Path.Combine(content.folderPath, knowledgeFileName),
                WrapInHtmlPage($"Knowledge - {content.filename}", body.ToString(), getNavigationHtml()));

            // Build individual knowledge detail pages
            Directory.CreateDirectory(Path.Combine(content.folderPath, "Knowledge"));
            foreach (BotComponent ks in knowledgeSources.OrderBy(k => k.Name))
            {
                buildKnowledgeDetailPage(ks);
            }
            foreach (BotComponent fk in fileKnowledge.OrderBy(k => k.Name))
            {
                buildKnowledgeDetailPage(fk);
            }
        }

        private void buildKnowledgeDetailPage(BotComponent knowledge)
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable(true));
            body.AppendLine(Heading(2, knowledge.Name));

            // Properties table
            body.Append(TableStart("Property", "Value"));
            body.Append(TableRow("Name", knowledge.Name));
            if (knowledge.ComponentType == 16)
            {
                body.Append(TableRow("Source Type", knowledge.GetSourceKindDisplayName()));
                string officialSource = knowledge.GetOfficialSourceDisplayName();
                if (!string.IsNullOrEmpty(officialSource))
                    body.Append(TableRow("Official Source", officialSource));
                string site = knowledge.GetKnowledgeSourceSite();
                if (!string.IsNullOrEmpty(site))
                    body.Append(TableRowRaw("URL", Link(site, site)));
            }
            else if (knowledge.ComponentType == 14)
            {
                string mimeType = !string.IsNullOrEmpty(knowledge.FileDataMimeType) ? $" ({knowledge.FileDataMimeType})" : "";
                body.Append(TableRow("Source Type", "File" + mimeType));
                if (!string.IsNullOrEmpty(knowledge.FileDataName))
                    body.Append(TableRow("File Name", knowledge.FileDataName));
            }
            body.AppendLine(TableEnd());

            // Description
            if (!string.IsNullOrEmpty(knowledge.Description))
            {
                body.AppendLine(Heading(3, "Description"));
                body.AppendLine(ParagraphWithLinebreaks(knowledge.Description));
            }

            // Dataverse-specific: Selected Tables and Synonyms
            if (knowledge.ComponentType == 16 && knowledge.GetSourceKindDisplayName() == "Dataverse")
            {
                var tables = content.agent.GetDataverseTablesForKnowledge(knowledge);
                if (tables.Count > 0)
                {
                    body.AppendLine(Heading(3, "Selected Tables"));
                    bool canLinkSolutionKnowledge = content.context?.Config?.documentSolution == true && content.context?.Solution != null;
                    string solutionHtmlKnowledge = canLinkSolutionKnowledge ? CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName) : null;
                    body.Append(TableStart("Table Name", "Logical Name"));
                    foreach (var table in tables.OrderBy(t => t.Name))
                    {
                        if (canLinkSolutionKnowledge)
                        {
                            string anchor = CrossDocLinkHelper.GetSolutionTableHtmlAnchor(table.EntityLogicalName);
                            body.Append(TableRowRaw(Link(table.Name, "../../" + solutionHtmlKnowledge + anchor), Encode(table.EntityLogicalName)));
                        }
                        else
                        {
                            body.Append(TableRow(table.Name, table.EntityLogicalName));
                        }
                    }
                    body.AppendLine(TableEnd());

                    // Synonyms/Glossary per table
                    foreach (var table in tables.OrderBy(t => t.Name))
                    {
                        var synonyms = content.agent.GetSynonymsForEntity(table);
                        if (synonyms.Count > 0)
                        {
                            body.AppendLine(Heading(3, $"Glossary: {table.Name}"));
                            body.Append(TableStart("Column", "Description"));
                            foreach (var syn in synonyms.OrderBy(s => s.ColumnLogicalName))
                            {
                                body.Append(TableRow(syn.ColumnLogicalName, syn.Description ?? ""));
                            }
                            body.AppendLine(TableEnd());
                        }
                    }
                }
            }

            string detailFile = knowledgeDetailFileNames.GetValueOrDefault(knowledge.Name, knowledge.Name + ".html");
            SaveHtmlFile(Path.Combine(content.folderPath, "Knowledge", detailFile),
                WrapInHtmlPage($"Knowledge: {knowledge.Name}", body.ToString(), getNavigationHtml(true), "../style.css"));
        }

        private void addAgentTools()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.Tools));

            var tools = content.GetResolvedToolInfos();
            if (tools.Count == 0)
            {
                body.AppendLine(Paragraph("No tools configured."));
            }
            else
            {
                // Summary table matching Copilot Studio UI columns
                body.Append(TableStart("Name", "Type", "Available to", "Trigger", "Enabled"));
                foreach (AgentToolInfo tool in tools)
                {
                    body.Append(TableRow(tool.Name, tool.ToolType, tool.AvailableTo ?? "", tool.Trigger ?? "", tool.Enabled ? "On" : "Off"));
                }
                body.AppendLine(TableEnd());

                // Detail per tool
                foreach (AgentToolInfo tool in tools)
                {
                    body.AppendLine(HeadingWithId(3, tool.Name, SanitizeAnchorId(tool.Name)));

                    // Details section
                    body.AppendLine(Heading(4, "Details"));
                    body.Append(TableStart("Property", "Value"));
                    body.Append(TableRow("Name", tool.Name));
                    if (!string.IsNullOrEmpty(tool.Description))
                        body.Append(TableRow("Description", tool.Description));
                    body.Append(TableRow("Type", tool.ToolType));
                    body.Append(TableRow("Available to", tool.AvailableTo ?? ""));
                    body.Append(TableRow("Trigger", tool.Trigger ?? ""));
                    body.Append(TableRow("Enabled", tool.Enabled ? "On" : "Off"));
                    if (!string.IsNullOrEmpty(tool.ConnectionReference))
                        body.Append(TableRow("Connection Reference", tool.ConnectionReference));
                    if (!string.IsNullOrEmpty(tool.OperationId))
                        body.Append(TableRow("Operation", tool.OperationId));
                    if (!string.IsNullOrEmpty(tool.FlowId))
                    {
                        if (content.context?.Config?.documentFlows == true)
                        {
                            FlowEntity flow = content.context.GetFlowById(tool.FlowId);
                            if (flow != null)
                            {
                                string href = "../" + CrossDocLinkHelper.GetFlowDocHtmlPath(flow.Name);
                                body.Append(TableRowRaw("Flow ID", Link(tool.FlowId, href)));
                            }
                            else
                            {
                                body.Append(TableRow("Flow ID", tool.FlowId));
                            }
                        }
                        else
                        {
                            body.Append(TableRow("Flow ID", tool.FlowId));
                        }
                    }
                    if (!string.IsNullOrEmpty(tool.AgentFlowName))
                    {
                        if (content.context?.Config?.documentFlows == true)
                        {
                            FlowEntity flow = content.context.GetFlowById(tool.FlowId);
                            if (flow != null)
                            {
                                string href = "../" + CrossDocLinkHelper.GetFlowDocHtmlPath(flow.Name);
                                body.Append(TableRowRaw("Agent Flow", Link(tool.AgentFlowName, href)));
                            }
                            else
                            {
                                body.Append(TableRow("Agent Flow", tool.AgentFlowName));
                            }
                        }
                        else
                        {
                            body.Append(TableRow("Agent Flow", tool.AgentFlowName));
                        }
                    }
                    if (!string.IsNullOrEmpty(tool.ModelParameters))
                        body.Append(TableRow("Model Parameters", tool.ModelParameters));
                    if (!string.IsNullOrEmpty(tool.OperationDetailsKind))
                        body.Append(TableRow("Protocol", tool.OperationDetailsKind));
                    if (!string.IsNullOrEmpty(tool.ConnectionMode))
                        body.Append(TableRow("Connection Mode", tool.ConnectionMode));
                    if (!string.IsNullOrEmpty(tool.ToolFilterKind))
                    {
                        string filterLabel = tool.ToolFilterKind == "UseSpecificTools" ? "Specific Tools" : tool.ToolFilterKind == "UseAllTools" ? "All Tools" : tool.ToolFilterKind;
                        body.Append(TableRow("Tool Filter", filterLabel));
                    }
                    body.AppendLine(TableEnd());

                    if (tool.AllowedTools.Count > 0)
                    {
                        body.AppendLine(Heading(4, "Allowed Tools"));
                        body.Append(TableStart("#", "Tool Name"));
                        int toolIndex = 1;
                        foreach (string t in tool.AllowedTools.OrderBy(t => t))
                        {
                            body.Append(TableRow(toolIndex.ToString(), t));
                            toolIndex++;
                        }
                        body.AppendLine(TableEnd());
                    }

                    // Connector / OpenAPI specification
                    if (tool.Connector != null)
                    {
                        body.AppendLine(Heading(4, "Connector"));
                        body.Append(TableStart("Property", "Value"));
                        if (!string.IsNullOrEmpty(tool.Connector.IconBlobBase64))
                            body.Append(TableRowRaw("Icon", $"<img src=\"data:image/png;base64,{tool.Connector.IconBlobBase64}\" alt=\"Connector Icon\" style=\"height:32px;\" />"));
                        body.Append(TableRow("Display Name", tool.Connector.DisplayName ?? ""));
                        if (!string.IsNullOrEmpty(tool.Connector.Description))
                            body.Append(TableRow("Description", tool.Connector.Description));
                        body.Append(TableRow("Connector Type", tool.Connector.ConnectorType == 1 ? "Custom" : tool.Connector.ConnectorType.ToString()));
                        if (!string.IsNullOrEmpty(tool.Connector.IconBrandColor))
                            body.Append(TableRowRaw("Brand Color", $"<span style=\"display:inline-block;width:16px;height:16px;background:{tool.Connector.IconBrandColor};border:1px solid #ccc;vertical-align:middle;\"></span> {System.Net.WebUtility.HtmlEncode(tool.Connector.IconBrandColor)}"));
                        body.AppendLine(TableEnd());

                        if (!string.IsNullOrEmpty(tool.Connector.OpenApiDefinitionJson))
                        {
                            body.AppendLine(Heading(5, "OpenAPI Definition"));
                            try
                            {
                                string formatted = Newtonsoft.Json.Linq.JToken.Parse(tool.Connector.OpenApiDefinitionJson).ToString(Newtonsoft.Json.Formatting.Indented);
                                body.AppendLine(PreCodeBlock(formatted));
                            }
                            catch
                            {
                                body.AppendLine(PreCodeBlock(tool.Connector.OpenApiDefinitionJson));
                            }
                        }

                        if (!string.IsNullOrEmpty(tool.Connector.ConnectionParametersJson))
                        {
                            body.AppendLine(Heading(5, "Connection Parameters"));
                            try
                            {
                                string formatted = Newtonsoft.Json.Linq.JToken.Parse(tool.Connector.ConnectionParametersJson).ToString(Newtonsoft.Json.Formatting.Indented);
                                body.AppendLine(PreCodeBlock(formatted));
                            }
                            catch
                            {
                                body.AppendLine(PreCodeBlock(tool.Connector.ConnectionParametersJson));
                            }
                        }

                        if (!string.IsNullOrEmpty(tool.Connector.PolicyTemplateInstancesJson))
                        {
                            body.AppendLine(Heading(5, "Policy Template Instances"));
                            try
                            {
                                string formatted = Newtonsoft.Json.Linq.JToken.Parse(tool.Connector.PolicyTemplateInstancesJson).ToString(Newtonsoft.Json.Formatting.Indented);
                                body.AppendLine(PreCodeBlock(formatted));
                            }
                            catch
                            {
                                body.AppendLine(PreCodeBlock(tool.Connector.PolicyTemplateInstancesJson));
                            }
                        }
                    }

                    // Inputs
                    if (tool.Inputs.Count > 0)
                    {
                        body.AppendLine(Heading(4, "Inputs"));
                        body.Append(TableStart("Input name", "Fill using", "Type", "Description"));
                        foreach (var input in tool.Inputs)
                        {
                            body.Append(TableRow(
                                input.Name + (input.IsRequired ? " *" : ""),
                                input.FillUsing ?? "",
                                input.DataType ?? "",
                                input.Description ?? ""));
                        }
                        body.AppendLine(TableEnd());
                    }

                    // Outputs
                    if (tool.Outputs.Count > 0)
                    {
                        body.AppendLine(Heading(4, "Outputs"));
                        body.Append(TableStart("Output name", "Type", "Description"));
                        foreach (var output in tool.Outputs)
                        {
                            body.Append(TableRow(output.Name, output.DataType ?? "", output.Description ?? ""));
                        }
                        body.AppendLine(TableEnd());
                    }

                    // Completion
                    if (!string.IsNullOrEmpty(tool.ResponseActivity) || !string.IsNullOrEmpty(tool.OutputMode))
                    {
                        body.AppendLine(Heading(4, "Completion"));
                        body.Append(TableStart("Property", "Value"));
                        if (!string.IsNullOrEmpty(tool.ResponseActivity))
                            body.Append(TableRow("After running", "Send specific response"));
                        if (!string.IsNullOrEmpty(tool.ResponseMode))
                            body.Append(TableRow("Response Mode", tool.ResponseMode));
                        if (!string.IsNullOrEmpty(tool.OutputMode))
                            body.Append(TableRow("Output Mode", tool.OutputMode));
                        body.AppendLine(TableEnd());
                        if (!string.IsNullOrEmpty(tool.ResponseActivity))
                        {
                            body.AppendLine(Heading(5, "Message to display"));
                            body.AppendLine(Paragraph(tool.ResponseActivity));
                        }
                    }

                    // Prompt text (for prompt tools)
                    if (!string.IsNullOrEmpty(tool.PromptText))
                    {
                        body.AppendLine(Heading(4, "Prompt"));
                        body.AppendLine(Paragraph(tool.PromptText));
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

                // Global Variable Usage Tracking
                var usageMap = content.GetGlobalVariableUsageMap();
                if (usageMap.Count > 0)
                {
                    body.AppendLine(Heading(3, "Global Variable Usage"));
                    body.AppendLine(Paragraph("Cross-reference showing which topics read or write each global variable."));

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
                        body.AppendLine(Heading(4, header));

                        body.Append(TableStart("Topic", "Access", "Context"));
                        foreach (var entry in usageMap[varName].OrderBy(e => e.AccessType).ThenBy(e => e.TopicName))
                        {
                            string topicCell = topicFileNames.TryGetValue(entry.TopicName, out var tf) ? Link(entry.TopicName, "Topics/" + tf) : Encode(entry.TopicName);
                            body.Append(TableRowRaw(topicCell, Encode(entry.AccessType.ToString()), Encode(entry.Context)));
                        }
                        body.AppendLine(TableEnd());
                    }
                }
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
            var agentInfos = content.GetResolvedConnectedAgentInfos();
            if (agentInfos.Count > 0)
            {
                body.Append(TableStart("Name", "Connection Type", "Description", "History Type"));
                foreach (var agentInfo in agentInfos.OrderBy(a => a.Name))
                {
                    string safeName = CharsetHelper.GetSafeName(agentInfo.Name);
                    string agentFolder = "AgentDoc " + safeName;
                    string agentFile = ("index-" + safeName + ".html").Replace(" ", "-");
                    string agentLink = "../" + agentFolder + "/" + agentFile;
                    body.Append(TableRowRaw(Link(agentInfo.Name, agentLink), Encode(agentInfo.ConnectionType), Encode(agentInfo.Description), Encode(agentInfo.HistoryType)));
                }
                body.AppendLine(TableEnd());
            }
            else
            {
                body.AppendLine(Paragraph("No connected agents configured."));
            }
            SaveHtmlFile(Path.Combine(content.folderPath, agentsFileName),
                WrapInHtmlPage($"Agents - {content.filename}", body.ToString(), getNavigationHtml()));
        }

        private void addAgentTopics()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, $"Agent - {content.filename}"));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.Topics));

            // Topic Data Flow diagram
            string dataFlowFile = "topic-dataflow.svg";
            if (File.Exists(Path.Combine(content.folderPath, dataFlowFile)))
            {
                body.AppendLine(Heading(3, "Topic Data Flow"));
                body.AppendLine(Paragraph("Visualization of how topics call each other via BeginDialog/ReplaceDialog, with data passed between them."));
                body.AppendLine(ParagraphRaw($"<a href=\"{Encode(dataFlowFile)}\" target=\"_blank\">{Image("Topic Data Flow Diagram", dataFlowFile)}</a>"));
                body.AppendLine("<br/>");

                var calls = content.GetTopicDataFlowInfo();
                if (calls.Count > 0)
                {
                    body.Append(TableStart("Source Topic", "Target Topic", "Call Type", "Data Passed"));
                    foreach (var call in calls)
                    {
                        string dataPassed = call.InputBindings.Count > 0 ? string.Join(", ", call.InputBindings.Keys) : "";
                        string sourceCell = topicFileNames.TryGetValue(call.SourceTopicName, out var sf) ? Link(call.SourceTopicName, "Topics/" + sf) : Encode(call.SourceTopicName);
                        string targetCell = topicFileNames.TryGetValue(call.TargetTopicName, out var tf) ? Link(call.TargetTopicName, "Topics/" + tf) : Encode(call.TargetTopicName);
                        body.Append(TableRowRaw(sourceCell, targetCell, Encode(call.CallKind), Encode(dataPassed)));
                    }
                    body.AppendLine(TableEnd());
                }
            }

            body.AppendLine(Heading(3, "Topic List"));
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

            // Variables (with Scope column)
            var variables = topic.GetTopicVariables();
            if (variables.Count > 0)
            {
                topicBody.AppendLine(Heading(3, "Variables"));
                topicBody.Append(TableStart("Variable", "Scope", "Context"));
                foreach (var (variable, context) in variables)
                {
                    string scope = variable.StartsWith("Global.") ? "Global"
                        : variable.StartsWith("System.") ? "System"
                        : "Topic";
                    topicBody.Append(TableRow(variable, scope, context));
                }
                topicBody.AppendLine(TableEnd());
            }

            // Topic flow diagram
            string graphFile = topic.getTopicFileName() + "-detailed.svg";
            if (File.Exists(Path.Combine(content.folderPath, "Topics", graphFile)))
            {
                topicBody.AppendLine(Heading(3, "Topic Flow"));
                topicBody.AppendLine(ParagraphRaw($"<a href=\"{Encode(graphFile)}\" target=\"_blank\">{Image("Topic Flow Diagram", graphFile)}</a>"));
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
            body.Append(TableRow("Response Model", content.agent.GetResponseModelDisplayName()));
            var settingsModelHint = content.agent.GetModelNameHint();
            if (!string.IsNullOrEmpty(settingsModelHint))
                body.Append(TableRow("Model Name Hint", settingsModelHint));
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
