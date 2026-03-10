using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using PowerDocu.Common;
using Grynwald.MarkdownGenerator;
using Svg;

namespace PowerDocu.AgentDocumenter
{
    class AgentMarkdownBuilder : MarkdownBuilder
    {
        private readonly AgentDocumentationContent content;
        private readonly string mainDocumentFileName, knowledgeFileName, toolsFileName, agentsFileName, topicsFileName, channelsFileName, settingsFileName, entitiesFileName, variablesFileName;
        private readonly MdDocument mainDocument, knowledgeDocument, toolsDocument, agentsDocument, topicsDocument, channelsDocument, settingsDocument, entitiesDocument, variablesDocument;
        private readonly Dictionary<string, MdDocument> topicsDocuments = new Dictionary<string, MdDocument>();
        private readonly Dictionary<string, MdDocument> knowledgeDetailDocuments = new Dictionary<string, MdDocument>();
        private readonly DocumentSet<MdDocument> set;
        private MdTable metadataTable;

        public AgentMarkdownBuilder(AgentDocumentationContent contentdocumentation)
        {
            content = contentdocumentation;
            Directory.CreateDirectory(content.folderPath);
            mainDocumentFileName = ("index " + content.filename + ".md").Replace(" ", "-");
            knowledgeFileName = ("knowledge " + content.filename + ".md").Replace(" ", "-");
            toolsFileName = ("tools " + content.filename + ".md").Replace(" ", "-");
            agentsFileName = ("agents " + content.filename + ".md").Replace(" ", "-");
            topicsFileName = ("topics " + content.filename + ".md").Replace(" ", "-");
            channelsFileName = ("channels " + content.filename + ".md").Replace(" ", "-");
            settingsFileName = ("settings " + content.filename + ".md").Replace(" ", "-");
            entitiesFileName = ("entities " + content.filename + ".md").Replace(" ", "-");
            variablesFileName = ("variables " + content.filename + ".md").Replace(" ", "-");
            set = new DocumentSet<MdDocument>();
            mainDocument = set.CreateMdDocument(mainDocumentFileName);
            knowledgeDocument = set.CreateMdDocument(knowledgeFileName);
            toolsDocument = set.CreateMdDocument(toolsFileName);
            agentsDocument = set.CreateMdDocument(agentsFileName);
            //a dedicated document for each topic
            topicsDocument = set.CreateMdDocument(topicsFileName);
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                topicsDocuments.Add(topic.getTopicFileName(), set.CreateMdDocument("Topics/" + ("topic " + topic.getTopicFileName() + " " + content.filename + ".md").Replace(" ", "-")));
            }
            foreach (BotComponent ks in content.agent.GetKnowledge().Concat(content.agent.GetFileKnowledge()).OrderBy(k => k.Name))
            {
                string key = CharsetHelper.GetSafeName(ks.Name);
                knowledgeDetailDocuments[key] = set.CreateMdDocument("Knowledge/" + ("knowledge " + key + " " + content.filename + ".md").Replace(" ", "-"));
            }
            channelsDocument = set.CreateMdDocument(channelsFileName);
            settingsDocument = set.CreateMdDocument(settingsFileName);
            entitiesDocument = set.CreateMdDocument(entitiesFileName);
            variablesDocument = set.CreateMdDocument(variablesFileName);
            addAgentOverview();
            addAgentKnowledgeInfo();
            addAgentTools();
            addAgentAgentsInfo();
            addAgentTopics();
            addAgentEntities();
            addAgentVariables();
            addAgentChannels();
            addAgentSettings();
            set.Save(content.folderPath);
            NotificationHelper.SendNotification("Created Markdown documentation for " + content.filename);
        }

        private static readonly string NotInExportMessage = "This setting is not available in the solution export.";

        private void addAgentSettings()
        {
            var config = content.agent.Configuration;
            var ai = config?.aISettings;

            // Generative AI
            settingsDocument.Root.Add(new MdHeading("Generative AI", 2));
            List<MdTableRow> genAiRows = new List<MdTableRow>
            {
                new MdTableRow("Generative Actions", config?.settings?.GenerativeActionsEnabled == true ? "Enabled" : "Disabled"),
                new MdTableRow("Use Model Knowledge", ai?.useModelKnowledge == true ? "Yes" : "No"),
                new MdTableRow("File Analysis", ai?.isFileAnalysisEnabled == true ? "Enabled" : "Disabled"),
                new MdTableRow("Semantic Search", ai?.isSemanticSearchEnabled == true ? "Enabled" : "Disabled"),
                new MdTableRow("Content Moderation", ai?.contentModeration ?? "Unknown"),
                new MdTableRow("Opt-in to Latest Models", ai?.optInUseLatestModels == true ? "Yes" : "No")
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), genAiRows));

            // Security
            settingsDocument.Root.Add(new MdHeading("Security", 2));
            List<MdTableRow> secRows = new List<MdTableRow>
            {
                new MdTableRow("Authentication Mode", content.agent.GetAuthenticationModeDisplayName()),
                new MdTableRow("Authentication Trigger", content.agent.GetAuthenticationTriggerDisplayName())
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), secRows));

            // Connection settings
            settingsDocument.Root.Add(new MdHeading("Connection settings", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Authoring canvas
            settingsDocument.Root.Add(new MdHeading("Authoring canvas", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Entities
            settingsDocument.Root.Add(new MdHeading("Entities", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Skills
            settingsDocument.Root.Add(new MdHeading("Skills", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Voice
            settingsDocument.Root.Add(new MdHeading("Voice", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Languages
            settingsDocument.Root.Add(new MdHeading("Languages", 2));
            List<MdTableRow> langRows = new List<MdTableRow>
            {
                new MdTableRow("Primary Language", content.agent.GetLanguageDisplayName())
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), langRows));

            // Language understanding
            settingsDocument.Root.Add(new MdHeading("Language understanding", 2));
            List<MdTableRow> luRows = new List<MdTableRow>
            {
                new MdTableRow("Recognizer", content.agent.GetRecognizerDisplayName())
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), luRows));

            // Component collections
            settingsDocument.Root.Add(new MdHeading("Component collections", 2));
            settingsDocument.Root.Add(new MdParagraph(new MdTextSpan(NotInExportMessage)));

            // Advanced
            settingsDocument.Root.Add(new MdHeading("Advanced", 2));
            List<MdTableRow> advRows = new List<MdTableRow>
            {
                new MdTableRow("Template", content.agent.Template ?? ""),
                new MdTableRow("Runtime Provider", content.agent.RuntimeProvider.ToString())
            };
            settingsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Setting", "Value" }), advRows));
        }

        private void addAgentOverview()
        {
            List<MdTableRow> tableRows = new List<MdTableRow>
            {
                new MdTableRow("Agent Name", content.agent.Name)
            };
            //todo where is the agent logo stored?
            if (!String.IsNullOrEmpty(content.agent.IconBase64))
            {
                Directory.CreateDirectory(content.folderPath);
                Bitmap agentLogo = ImageHelper.ConvertBase64ToBitmap(content.agent.IconBase64);

                agentLogo.Save(content.folderPath + $"agentlogo-{content.filename.Replace(" ", "-")}.png");

                tableRows.Add(new MdTableRow("Agent Logo", new MdImageSpan("Agent Logo", $"agentlogo-{content.filename.Replace(" ", "-")}.png")));
                agentLogo.Dispose();
            }

            tableRows.Add(new MdTableRow(content.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion()));
            metadataTable = new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), tableRows);
            // prepare the common sections for top-level documents
            foreach (MdDocument doc in new[] { mainDocument, knowledgeDocument, toolsDocument, agentsDocument, topicsDocument, channelsDocument, settingsDocument, entitiesDocument, variablesDocument })
            {
                doc.Root.Add(new MdHeading($"Agent - {content.filename}", 1));
                doc.Root.Add(metadataTable);
                doc.Root.Add(getNavigationLinks());
            }
            // prepare the common sections for topic documents (in Topics subfolder)
            foreach (var kvp in topicsDocuments)
            {
                kvp.Value.Root.Add(new MdHeading($"Agent - {content.filename}", 1));
                kvp.Value.Root.Add(metadataTable);
                kvp.Value.Root.Add(getNavigationLinks(false));
            }
            // prepare the common sections for knowledge detail documents (in Knowledge subfolder)
            foreach (var kvp in knowledgeDetailDocuments)
            {
                kvp.Value.Root.Add(new MdHeading($"Agent - {content.filename}", 1));
                kvp.Value.Root.Add(metadataTable);
                kvp.Value.Root.Add(getNavigationLinks(false));
            }

            mainDocument.Root.Add(new MdHeading(content.Details, 2));
            mainDocument.Root.Add(new MdHeading(content.Description, 3));
            AddParagraphsWithLinebreaks(mainDocument, content.agent.GetDescription());
            mainDocument.Root.Add(new MdHeading(content.Orchestration, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"{content.OrchestrationText} - {content.agent.GetOrchestration()}")));
            mainDocument.Root.Add(new MdHeading(content.ResponseModel, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"{content.agent.GetResponseModel()}")));
            mainDocument.Root.Add(new MdHeading(content.Instructions, 3));
            AddParagraphsWithLinebreaks(mainDocument, content.agent.GetInstructions());
            mainDocument.Root.Add(new MdHeading(content.Knowledge, 3));
            var knowledgeSources = content.agent.GetKnowledge();
            var fileKnowledge = content.agent.GetFileKnowledge();
            if (knowledgeSources.Count > 0 || fileKnowledge.Count > 0)
            {
                List<MdTableRow> ksRows = new List<MdTableRow>();
                foreach (BotComponent ks in knowledgeSources)
                {
                    string key = CharsetHelper.GetSafeName(ks.Name);
                    string detailLink = "Knowledge/" + ("knowledge " + key + " " + content.filename + ".md").Replace(" ", "-");
                    string details = content.agent.GetKnowledgeDetailsSummary(ks);
                    string site = ks.GetKnowledgeSourceSite();
                    MdSpan detailsCell = !string.IsNullOrEmpty(site) ? new MdLinkSpan(details, site) : (MdSpan)new MdTextSpan(details);
                    ksRows.Add(new MdTableRow(new MdLinkSpan(ks.Name, detailLink), ks.GetSourceKindDisplayName(), detailsCell));
                }
                foreach (BotComponent fk in fileKnowledge)
                {
                    string key = CharsetHelper.GetSafeName(fk.Name);
                    string detailLink = "Knowledge/" + ("knowledge " + key + " " + content.filename + ".md").Replace(" ", "-");
                    string mimeType = !string.IsNullOrEmpty(fk.FileDataMimeType) ? $" ({fk.FileDataMimeType})" : "";
                    ksRows.Add(new MdTableRow(new MdLinkSpan(fk.Name, detailLink), "File" + mimeType, fk.FileDataName ?? ""));
                }
                mainDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Name", "Source Type", "Details" }), ksRows));
            }
            else
            {
                mainDocument.Root.Add(new MdParagraph(new MdTextSpan("No knowledge sources configured.")));
            }
            mainDocument.Root.Add(new MdHeading(content.Tools, 3));
            var overviewTools = content.GetResolvedToolInfos();
            if (overviewTools.Count > 0)
            {
                List<MdListItem> toolsList = new List<MdListItem>();
                foreach (AgentToolInfo tool in overviewTools)
                {
                    toolsList.Add(new MdListItem($"{tool.Name} ({tool.ToolType})"));
                }
                mainDocument.Root.Add(new MdBulletList(toolsList));
            }
            else
            {
                mainDocument.Root.Add(new MdParagraph(new MdTextSpan("No tools configured.")));
            }
            mainDocument.Root.Add(new MdHeading(content.Triggers, 3));
            var overviewTriggers = content.agent.GetTriggers();
            if (overviewTriggers.Count > 0)
            {
                List<MdListItem> triggersList = new List<MdListItem>();
                foreach (BotComponent trigger in overviewTriggers.OrderBy(t => t.Name))
                {
                    var (triggerKind, flowId, connectionType) = trigger.GetTriggerDetails();
                    string triggerInfo = !string.IsNullOrEmpty(connectionType) ? $" ({connectionType})" : "";
                    triggersList.Add(new MdListItem(trigger.Name + triggerInfo));
                }
                mainDocument.Root.Add(new MdBulletList(triggersList));
            }
            else
            {
                mainDocument.Root.Add(new MdParagraph(new MdTextSpan("No triggers configured.")));
            }
            mainDocument.Root.Add(new MdHeading(content.Agents, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan("Sub-agents are not available in the solution export.")));
            mainDocument.Root.Add(new MdHeading(content.Topics, 3));
            List<MdListItem> topicsList = new List<MdListItem>();
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name))
            {
                topicsList.Add(new MdListItem(new MdLinkSpan(topic.Name, "Topics/" + ("topic " + topic.getTopicFileName() + " " + content.filename + ".md").Replace(" ", "-"))));
            }
            mainDocument.Root.Add(new MdBulletList(topicsList));
            mainDocument.Root.Add(new MdHeading(content.SuggestedPrompts, 3));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan(content.SuggestedPromptsText)));
            tableRows = new List<MdTableRow>();
            Dictionary<string, string> conversationStarters = content.agent.GetSuggestedPrompts();
            foreach (var kvp in conversationStarters.OrderBy(x => x.Key))
            {
                tableRows.Add(new MdTableRow(kvp.Key, kvp.Value));
            }
            mainDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Prompt Title", "Prompt" }), tableRows));
        }

        private MdBulletList getNavigationLinks(bool topLevel = true)
        {
            MdListItem[] navItems = new MdListItem[] {
                new MdListItem(new MdLinkSpan("Overview", topLevel ? mainDocumentFileName : "../" + mainDocumentFileName)),
                new MdListItem(new MdLinkSpan("Knowledge", topLevel ? knowledgeFileName : "../" + knowledgeFileName)),
                new MdListItem(new MdLinkSpan("Tools", topLevel ? toolsFileName : "../" + toolsFileName)),
                new MdListItem(new MdLinkSpan("Entities", topLevel ? entitiesFileName : "../" + entitiesFileName)),
                new MdListItem(new MdLinkSpan("Variables", topLevel ? variablesFileName : "../" + variablesFileName)),
                new MdListItem(new MdLinkSpan("Agents", topLevel ? agentsFileName : "../" + agentsFileName)),
                new MdListItem(new MdLinkSpan("Topics", topLevel ? topicsFileName : "../" + topicsFileName)),
                new MdListItem(new MdLinkSpan("Channels", topLevel ? channelsFileName : "../" + channelsFileName)),
                new MdListItem(new MdLinkSpan("Settings", topLevel ? settingsFileName : "../" + settingsFileName))
                };
            return new MdBulletList(navItems);
        }


        /*
            private void addAppDetails()
            {
                List<MdTableRow> tableRows = new List<MdTableRow>();
                knowledgeDocument.Root.Add(new MdHeading(content.appProperties.headerAppProperties, 2));
                foreach (Expression property in content.appProperties.appProperties)
                {
                    if (!content.appProperties.propertiesToSkip.Contains(property.expressionOperator))
                    {
                        tableRows.Add(new MdTableRow(property.expressionOperator, property.expressionOperands[0].ToString()));
                    }
                }
                if (tableRows.Count > 0)
                {
                    knowledgeDocument.Root.Add(new MdTable(new MdTableRow("App Property", "Value"), tableRows));
                }
                knowledgeDocument.Root.Add(new MdHeading(content.appProperties.headerAppPreviewFlags, 2));
                tableRows = new List<MdTableRow>();
                if (content.appProperties.appPreviewsFlagProperty != null)
                {
                    foreach (Expression flagProp in content.appProperties.appPreviewsFlagProperty.expressionOperands)
                    {
                        tableRows.Add(new MdTableRow(flagProp.expressionOperator, flagProp.expressionOperands[0].ToString()));
                    }
                    if (tableRows.Count > 0)
                    {
                        knowledgeDocument.Root.Add(new MdTable(new MdTableRow("Preview Flag", "Value"), tableRows));
                    }
                }
            }

    */
        private void addAgentKnowledgeInfo()
        {
            knowledgeDocument.Root.Add(new MdHeading(content.Knowledge, 2));
            knowledgeDocument.Root.Add(new MdParagraph(new MdTextSpan("Knowledge sources for this agent.")));

            var knowledgeSources = content.agent.GetKnowledge();
            var fileKnowledge = content.agent.GetFileKnowledge();

            if (knowledgeSources.Count > 0 || fileKnowledge.Count > 0)
            {
                List<MdTableRow> ksRows = new List<MdTableRow>();
                foreach (BotComponent ks in knowledgeSources)
                {
                    string key = CharsetHelper.GetSafeName(ks.Name);
                    string detailLink = "Knowledge/" + ("knowledge " + key + " " + content.filename + ".md").Replace(" ", "-");
                    string details = content.agent.GetKnowledgeDetailsSummary(ks);
                    string officialSource = ks.GetOfficialSourceDisplayName();
                    string descriptionPreview = !string.IsNullOrEmpty(ks.Description) && ks.Description.Length > 100
                        ? ks.Description.Substring(0, 100) + "..."
                        : ks.Description ?? "";
                    string site = ks.GetKnowledgeSourceSite();
                    MdSpan detailsCell = !string.IsNullOrEmpty(site) ? new MdLinkSpan(details, site) : (MdSpan)new MdTextSpan(details);
                    ksRows.Add(new MdTableRow(new MdLinkSpan(ks.Name, detailLink), ks.GetSourceKindDisplayName(), officialSource, detailsCell, descriptionPreview));
                }
                foreach (BotComponent fk in fileKnowledge)
                {
                    string key = CharsetHelper.GetSafeName(fk.Name);
                    string detailLink = "Knowledge/" + ("knowledge " + key + " " + content.filename + ".md").Replace(" ", "-");
                    string mimeType = !string.IsNullOrEmpty(fk.FileDataMimeType) ? $" ({fk.FileDataMimeType})" : "";
                    string descriptionPreview = !string.IsNullOrEmpty(fk.Description) && fk.Description.Length > 100
                        ? fk.Description.Substring(0, 100) + "..."
                        : fk.Description ?? "";
                    ksRows.Add(new MdTableRow(new MdLinkSpan(fk.Name, detailLink), "File" + mimeType, "", fk.FileDataName ?? "", descriptionPreview));
                }
                knowledgeDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Name", "Source Type", "Official Source", "Details", "Description" }), ksRows));
            }
            else
            {
                knowledgeDocument.Root.Add(new MdParagraph(new MdTextSpan("No knowledge sources configured.")));
            }

            // Build individual knowledge detail documents
            foreach (BotComponent ks in knowledgeSources.OrderBy(k => k.Name))
            {
                buildKnowledgeDetailDocument(ks);
            }
            foreach (BotComponent fk in fileKnowledge.OrderBy(k => k.Name))
            {
                buildKnowledgeDetailDocument(fk);
            }
        }

        private void buildKnowledgeDetailDocument(BotComponent knowledge)
        {
            string key = CharsetHelper.GetSafeName(knowledge.Name);
            if (!knowledgeDetailDocuments.TryGetValue(key, out MdDocument doc)) return;

            doc.Root.Add(new MdHeading(knowledge.Name, 2));

            // Properties table
            List<MdTableRow> propRows = new List<MdTableRow>();
            propRows.Add(new MdTableRow("Name", knowledge.Name));
            if (knowledge.ComponentType == 16)
            {
                propRows.Add(new MdTableRow("Source Type", knowledge.GetSourceKindDisplayName()));
                string officialSource = knowledge.GetOfficialSourceDisplayName();
                if (!string.IsNullOrEmpty(officialSource))
                    propRows.Add(new MdTableRow("Official Source", officialSource));
                string site = knowledge.GetKnowledgeSourceSite();
                if (!string.IsNullOrEmpty(site))
                    propRows.Add(new MdTableRow("URL", new MdLinkSpan(site, site)));
            }
            else if (knowledge.ComponentType == 14)
            {
                string mimeType = !string.IsNullOrEmpty(knowledge.FileDataMimeType) ? $" ({knowledge.FileDataMimeType})" : "";
                propRows.Add(new MdTableRow("Source Type", "File" + mimeType));
                if (!string.IsNullOrEmpty(knowledge.FileDataName))
                    propRows.Add(new MdTableRow("File Name", knowledge.FileDataName));
            }
            doc.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), propRows));

            // Description
            if (!string.IsNullOrEmpty(knowledge.Description))
            {
                doc.Root.Add(new MdHeading("Description", 3));
                AddParagraphsWithLinebreaks(doc, knowledge.Description);
            }

            // Dataverse-specific: Selected Tables and Synonyms
            if (knowledge.ComponentType == 16 && knowledge.GetSourceKindDisplayName() == "Dataverse")
            {
                var tables = content.agent.GetDataverseTablesForKnowledge(knowledge);
                if (tables.Count > 0)
                {
                    doc.Root.Add(new MdHeading("Selected Tables", 3));
                    bool canLinkSolutionKnowledge = content.context?.Config?.documentSolution == true && content.context?.Solution != null;
                    string solutionMdKnowledge = canLinkSolutionKnowledge ? CrossDocLinkHelper.GetSolutionDocMdPath(content.context.Solution.UniqueName) : null;
                    List<MdTableRow> tableRows = new List<MdTableRow>();
                    foreach (var table in tables.OrderBy(t => t.Name))
                    {
                        MdSpan nameCell;
                        if (canLinkSolutionKnowledge)
                        {
                            string anchor = CrossDocLinkHelper.GetSolutionTableMdAnchor(table.Name, table.EntityLogicalName);
                            nameCell = new MdLinkSpan(table.Name, "../../" + solutionMdKnowledge + anchor);
                        }
                        else
                        {
                            nameCell = new MdTextSpan(table.Name);
                        }
                        tableRows.Add(new MdTableRow(nameCell, new MdTextSpan(table.EntityLogicalName)));
                    }
                    doc.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Table Name", "Logical Name" }), tableRows));

                    // Synonyms/Glossary per table
                    foreach (var table in tables.OrderBy(t => t.Name))
                    {
                        var synonyms = content.agent.GetSynonymsForEntity(table);
                        if (synonyms.Count > 0)
                        {
                            doc.Root.Add(new MdHeading($"Glossary: {table.Name}", 3));
                            List<MdTableRow> synRows = new List<MdTableRow>();
                            foreach (var syn in synonyms.OrderBy(s => s.ColumnLogicalName))
                            {
                                synRows.Add(new MdTableRow(syn.ColumnLogicalName, syn.Description ?? ""));
                            }
                            doc.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Column", "Description" }), synRows));
                        }
                    }
                }
            }
        }

        private void addAgentTopics()
        {
            topicsDocument.Root.Add(new MdHeading(content.Topics, 2));
            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (BotComponent topic in content.agent.GetTopics().OrderBy(o => o.Name).ToList())
            {
                string topicType = topic.GetComponentTypeDisplayName();
                string triggerType = topic.GetTriggerTypeForTopic();
                tableRows.Add(new MdTableRow(
                    new MdLinkSpan(topic.Name, "Topics/" + ("topic " + topic.getTopicFileName() + " " + content.filename + ".md").Replace(" ", "-")),
                    topicType,
                    triggerType,
                    topic.GetTopicKind() == "KnowledgeSourceConfiguration" ? "Knowledge" : topic.GetTopicKind()));

                // Fill per-topic document
                topicsDocuments.TryGetValue(topic.getTopicFileName(), out MdDocument topicDoc);
                if (topicDoc != null)
                {
                    addTopicDetails(topicDoc, topic);
                }
            }
            topicsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Name", "Type", "Trigger", "Kind" }), tableRows));
        }

        private void addTopicDetails(MdDocument topicDoc, BotComponent topic)
        {
            topicDoc.Root.Add(new MdHeading("Topic: " + topic.Name, 2));

            // Metadata table
            List<MdTableRow> metaRows = new List<MdTableRow>
            {
                new MdTableRow("Name", topic.Name),
                new MdTableRow("Type", topic.GetComponentTypeDisplayName()),
                new MdTableRow("Trigger", topic.GetTriggerTypeForTopic()),
                new MdTableRow("Topic Kind", topic.GetTopicKind())
            };
            if (!string.IsNullOrEmpty(topic.Description))
            {
                metaRows.Add(new MdTableRow("Description", topic.Description));
            }
            string modelDesc = topic.GetModelDescription();
            if (!string.IsNullOrEmpty(modelDesc))
            {
                metaRows.Add(new MdTableRow("Model Description", modelDesc));
            }
            string startBehavior = topic.GetStartBehavior();
            if (!string.IsNullOrEmpty(startBehavior))
            {
                metaRows.Add(new MdTableRow("Start Behavior", startBehavior));
            }
            topicDoc.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), metaRows));

            // Trigger queries
            List<string> triggerQueries = topic.GetTriggerQueries();
            if (triggerQueries.Count > 0)
            {
                topicDoc.Root.Add(new MdHeading("Trigger Queries", 3));
                List<MdListItem> queryItems = new List<MdListItem>();
                foreach (string query in triggerQueries)
                {
                    queryItems.Add(new MdListItem(query));
                }
                topicDoc.Root.Add(new MdBulletList(queryItems));
            }

            // Knowledge source details for KnowledgeSourceConfiguration topics
            if (topic.GetTopicKind() == "KnowledgeSourceConfiguration")
            {
                var (sourceKind, skillConfig) = topic.GetKnowledgeSourceDetails();
                topicDoc.Root.Add(new MdHeading("Knowledge Source", 3));
                List<MdTableRow> ksRows = new List<MdTableRow>();
                if (!string.IsNullOrEmpty(sourceKind))
                    ksRows.Add(new MdTableRow("Source Kind", sourceKind));
                if (!string.IsNullOrEmpty(skillConfig))
                    ksRows.Add(new MdTableRow("Skill Configuration", skillConfig));
                if (ksRows.Count > 0)
                    topicDoc.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), ksRows));
            }

            // Variables
            var variables = topic.GetTopicVariables();
            if (variables.Count > 0)
            {
                topicDoc.Root.Add(new MdHeading("Variables", 3));
                List<MdTableRow> varRows = new List<MdTableRow>();
                foreach (var (variable, context) in variables)
                {
                    varRows.Add(new MdTableRow(variable, context));
                }
                topicDoc.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Variable", "Context" }), varRows));
            }

            // Topic flow diagram
            string graphFile = topic.getTopicFileName() + "-detailed.svg";
            if (File.Exists(Path.Combine(content.folderPath, "Topics", graphFile)))
            {
                topicDoc.Root.Add(new MdHeading("Topic Flow", 3));
                topicDoc.Root.Add(new MdParagraph(new MdRawMarkdownSpan($"[![Topic Flow Diagram]({graphFile})]({graphFile})")));
            }
        }

        private void addAgentChannels()
        {
            channelsDocument.Root.Add(new MdHeading("Channels", 2));
            channelsDocument.Root.Add(new MdParagraph(new MdTextSpan("Channels are not exported with the solution and are not documented automatically.")));
        }

        private void addAgentTools()
        {
            toolsDocument.Root.Add(new MdHeading(content.Tools, 2));

            var tools = content.GetResolvedToolInfos();
            if (tools.Count == 0)
            {
                toolsDocument.Root.Add(new MdParagraph(new MdTextSpan("No tools configured.")));
                return;
            }

            // Summary table matching Copilot Studio UI columns
            List<MdTableRow> summaryRows = new List<MdTableRow>();
            foreach (AgentToolInfo tool in tools)
            {
                summaryRows.Add(new MdTableRow(
                    tool.Name,
                    tool.ToolType,
                    tool.AvailableTo ?? "",
                    tool.Trigger ?? "",
                    tool.Enabled ? "On" : "Off"));
            }
            toolsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Name", "Type", "Available to", "Trigger", "Enabled" }), summaryRows));

            // Detail per tool
            foreach (AgentToolInfo tool in tools)
            {
                toolsDocument.Root.Add(new MdHeading(tool.Name, 3));

                // Details section
                toolsDocument.Root.Add(new MdHeading("Details", 4));
                List<MdTableRow> detailRows = new List<MdTableRow>();
                detailRows.Add(new MdTableRow("Name", tool.Name));
                if (!string.IsNullOrEmpty(tool.Description))
                    detailRows.Add(new MdTableRow("Description", tool.Description));
                detailRows.Add(new MdTableRow("Type", tool.ToolType));
                detailRows.Add(new MdTableRow("Available to", tool.AvailableTo ?? ""));
                detailRows.Add(new MdTableRow("Trigger", tool.Trigger ?? ""));
                detailRows.Add(new MdTableRow("Enabled", tool.Enabled ? "On" : "Off"));
                if (!string.IsNullOrEmpty(tool.ConnectionReference))
                    detailRows.Add(new MdTableRow("Connection Reference", tool.ConnectionReference));
                if (!string.IsNullOrEmpty(tool.OperationId))
                    detailRows.Add(new MdTableRow("Operation", tool.OperationId));
                if (!string.IsNullOrEmpty(tool.FlowId))
                {
                    if (content.context?.Config?.documentFlows == true)
                    {
                        FlowEntity flow = content.context.GetFlowById(tool.FlowId);
                        if (flow != null)
                        {
                            string href = "../" + CrossDocLinkHelper.GetFlowDocMdPath(flow.Name);
                            detailRows.Add(new MdTableRow(new MdTextSpan("Flow ID"), new MdLinkSpan(tool.FlowId, href)));
                        }
                        else
                        {
                            detailRows.Add(new MdTableRow("Flow ID", tool.FlowId));
                        }
                    }
                    else
                    {
                        detailRows.Add(new MdTableRow("Flow ID", tool.FlowId));
                    }
                }
                if (!string.IsNullOrEmpty(tool.AgentFlowName))
                {
                    if (content.context?.Config?.documentFlows == true)
                    {
                        FlowEntity flow = content.context.GetFlowById(tool.FlowId);
                        if (flow != null)
                        {
                            string href = "../" + CrossDocLinkHelper.GetFlowDocMdPath(flow.Name);
                            detailRows.Add(new MdTableRow(new MdTextSpan("Agent Flow"), new MdLinkSpan(tool.AgentFlowName, href)));
                        }
                        else
                        {
                            detailRows.Add(new MdTableRow("Agent Flow", tool.AgentFlowName));
                        }
                    }
                    else
                    {
                        detailRows.Add(new MdTableRow("Agent Flow", tool.AgentFlowName));
                    }
                }
                if (!string.IsNullOrEmpty(tool.ModelParameters))
                    detailRows.Add(new MdTableRow("Model Parameters", tool.ModelParameters));
                toolsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), detailRows));

                // Inputs section
                if (tool.Inputs.Count > 0)
                {
                    toolsDocument.Root.Add(new MdHeading("Inputs", 4));
                    List<MdTableRow> inputRows = new List<MdTableRow>();
                    foreach (var input in tool.Inputs)
                    {
                        inputRows.Add(new MdTableRow(
                            input.Name + (input.IsRequired ? " *" : ""),
                            input.FillUsing ?? "",
                            input.DataType ?? "",
                            input.Description ?? ""));
                    }
                    toolsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Input name", "Fill using", "Type", "Description" }), inputRows));
                }

                // Outputs section
                if (tool.Outputs.Count > 0)
                {
                    toolsDocument.Root.Add(new MdHeading("Outputs", 4));
                    List<MdTableRow> outputRows = new List<MdTableRow>();
                    foreach (var output in tool.Outputs)
                    {
                        outputRows.Add(new MdTableRow(
                            output.Name,
                            output.DataType ?? "",
                            output.Description ?? ""));
                    }
                    toolsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Output name", "Type", "Description" }), outputRows));
                }

                // Completion section
                if (!string.IsNullOrEmpty(tool.ResponseActivity) || !string.IsNullOrEmpty(tool.OutputMode))
                {
                    toolsDocument.Root.Add(new MdHeading("Completion", 4));
                    List<MdTableRow> completionRows = new List<MdTableRow>();
                    if (!string.IsNullOrEmpty(tool.ResponseActivity))
                        completionRows.Add(new MdTableRow("After running", "Send specific response"));
                    if (!string.IsNullOrEmpty(tool.ResponseMode))
                        completionRows.Add(new MdTableRow("Response Mode", tool.ResponseMode));
                    if (!string.IsNullOrEmpty(tool.OutputMode))
                        completionRows.Add(new MdTableRow("Output Mode", tool.OutputMode));
                    toolsDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), completionRows));
                    if (!string.IsNullOrEmpty(tool.ResponseActivity))
                    {
                        toolsDocument.Root.Add(new MdHeading("Message to display", 5));
                        AddParagraphsWithLinebreaks(toolsDocument, tool.ResponseActivity);
                    }
                }

                // Prompt text section (for prompt tools)
                if (!string.IsNullOrEmpty(tool.PromptText))
                {
                    toolsDocument.Root.Add(new MdHeading("Prompt", 4));
                    AddParagraphsWithLinebreaks(toolsDocument, tool.PromptText);
                }
            }
        }

        private void addAgentEntities()
        {
            entitiesDocument.Root.Add(new MdHeading(content.Entities, 2));

            var entities = content.agent.GetEntities();
            if (entities.Count == 0)
            {
                entitiesDocument.Root.Add(new MdParagraph(new MdTextSpan("No custom entities defined.")));
                return;
            }

            foreach (BotComponent entity in entities.OrderBy(e => e.Name))
            {
                string entityKind = entity.GetTopicKind();
                entitiesDocument.Root.Add(new MdHeading(entity.Name, 3));

                List<MdTableRow> metaRows = new List<MdTableRow>
                {
                    new MdTableRow("Name", entity.Name),
                    new MdTableRow("Kind", entityKind)
                };
                if (!string.IsNullOrEmpty(entity.Description))
                    metaRows.Add(new MdTableRow("Description", entity.Description));

                if (entityKind == "RegexEntity")
                {
                    string pattern = entity.GetEntityPattern();
                    if (!string.IsNullOrEmpty(pattern))
                        metaRows.Add(new MdTableRow("Pattern", $"`{pattern}`"));
                }
                entitiesDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Property", "Value" }), metaRows));

                if (entityKind == "ClosedListEntity")
                {
                    var items = entity.GetEntityItems();
                    if (items.Count > 0)
                    {
                        entitiesDocument.Root.Add(new MdHeading("Items", 4));
                        List<MdTableRow> itemRows = new List<MdTableRow>();
                        foreach (var (id, displayName) in items)
                        {
                            itemRows.Add(new MdTableRow(displayName, id));
                        }
                        entitiesDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Display Name", "ID" }), itemRows));
                    }
                }
            }
        }

        private void addAgentVariables()
        {
            variablesDocument.Root.Add(new MdHeading(content.Variables, 2));

            var variables = content.agent.GetVariables();
            if (variables.Count == 0)
            {
                variablesDocument.Root.Add(new MdParagraph(new MdTextSpan("No global variables defined.")));
                return;
            }

            List<MdTableRow> varRows = new List<MdTableRow>();
            foreach (BotComponent variable in variables.OrderBy(v => v.Name))
            {
                var (scope, aiVisibility, dataType, isExternalInit) = variable.GetVariableDetails();
                varRows.Add(new MdTableRow(
                    variable.Name,
                    scope,
                    dataType,
                    aiVisibility,
                    isExternalInit ? "Yes" : "No",
                    variable.Description ?? ""));
            }
            variablesDocument.Root.Add(new MdTable(new MdTableRow(new List<string>() { "Name", "Scope", "Data Type", "AI Visibility", "External Init", "Description" }), varRows));
        }

        private void addAgentAgentsInfo()
        {
            agentsDocument.Root.Add(new MdHeading(content.Agents, 2));
            agentsDocument.Root.Add(new MdParagraph(new MdTextSpan("Sub-agents are not available in the solution export.")));
        }

        private void AddParagraphsWithLinebreaks(MdDocument document, string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                document.Root.Add(new MdParagraph(new MdTextSpan("")));
                return;
            }
            string[] lines = text.Replace("\r\n", "\n").Split('\n');
            foreach (string line in lines)
            {
                document.Root.Add(new MdParagraph(new MdTextSpan(line)));
            }
        }
    }
}
