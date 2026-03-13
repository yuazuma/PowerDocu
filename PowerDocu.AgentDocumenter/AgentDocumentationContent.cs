using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerDocu.Common;

namespace PowerDocu.AgentDocumenter
{

    class AgentDocumentationContent
    {
        public string folderPath,
           filename;
        public AgentEntity agent;
        public DocumentationContext context;

        public string headerDocumentationGenerated = "Documentation generated at";
        public string Details = "Details";
        public string Description = "Description";
        public string Orchestration = "Orchestration";
        public string OrchestrationText = "Use generative AI to determine how best to respond to users and events.";
        public string ResponseModel = "Response model";
        public string Instructions = "Instructions";
        public string Knowledge = "Knowledge";
        public string WebSearch = "Web Search";
        public string Tools = "Tools";
        public string Triggers = "Triggers";
        public string Agents = "Agents";
        public string Topics = "Topics";
        public string Entities = "Entities";
        public string Variables = "Variables";
        public string SuggestedPrompts = "Suggested prompts";
        public string SuggestedPromptsText = "Suggest ways of starting conversations for Teams and Microsoft 365 channels.";

        public AgentDocumentationContent(AgentEntity agent, string path, DocumentationContext context = null)
        {
            NotificationHelper.SendNotification("Preparing documentation content for " + agent.Name);
            folderPath = path + CharsetHelper.GetSafeName(@"\AgentDoc " + agent.Name + @"\");
            Directory.CreateDirectory(folderPath);
            filename = CharsetHelper.GetSafeName(agent.Name);
            this.agent = agent;
            this.context = context;
        }

        /// <summary>
        /// Resolves a flow ID to its display name using the DocumentationContext.
        /// </summary>
        public string GetFlowNameForId(string flowId)
        {
            if (string.IsNullOrEmpty(flowId)) return flowId;
            return context?.GetFlowNameById(flowId) ?? flowId;
        }

        /// <summary>
        /// Resolves a Dataverse table schema name to its display name using the DocumentationContext.
        /// </summary>
        public string GetTableDisplayName(string schemaName)
        {
            if (string.IsNullOrEmpty(schemaName)) return schemaName;
            return context?.GetTableDisplayName(schemaName) ?? schemaName;
        }

        /// <summary>
        /// Returns all tool infos with flow names resolved via the DocumentationContext.
        /// </summary>
        public List<AgentToolInfo> GetResolvedToolInfos()
        {
            var tools = agent.GetAllToolInfos();
            if (context != null)
            {
                foreach (var tool in tools)
                {
                    if (!string.IsNullOrEmpty(tool.FlowId) && string.IsNullOrEmpty(tool.AgentFlowName))
                    {
                        tool.AgentFlowName = context.GetFlowNameById(tool.FlowId);
                    }
                }
            }
            return tools;
        }

        /// <summary>
        /// Returns all connected agent infos with names resolved via the DocumentationContext.
        /// For parent agents: lists connected child agents + other solution agents.
        /// For child agents: lists only the parent agent(s) that reference this agent.
        /// </summary>
        public List<ConnectedAgentInfo> GetResolvedConnectedAgentInfos()
        {
            // Start with explicitly connected agents (InvokeConnectedAgentTaskAction)
            var result = agent.GetAllConnectedAgentInfos();
            if (context != null)
            {
                // Resolve display names for connected agents
                foreach (var info in result)
                {
                    if (!string.IsNullOrEmpty(info.BotSchemaName))
                    {
                        string resolvedName = context.GetAgentNameBySchemaName(info.BotSchemaName);
                        if (resolvedName != info.BotSchemaName)
                            info.Name = resolvedName;
                    }
                }

                if (result.Count > 0)
                {
                    // This agent has connected agents — it's a parent/orchestrator.
                    // Also add other solution agents not already listed as connected.
                    var connectedSchemaNames = new HashSet<string>(
                        result.Where(r => !string.IsNullOrEmpty(r.BotSchemaName)).Select(r => r.BotSchemaName),
                        StringComparer.OrdinalIgnoreCase);

                    foreach (var otherAgent in context.Agents)
                    {
                        if (otherAgent.SchemaName.Equals(agent.SchemaName, StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (connectedSchemaNames.Contains(otherAgent.SchemaName))
                            continue;
                        result.Add(new ConnectedAgentInfo
                        {
                            Name = otherAgent.Name,
                            BotSchemaName = otherAgent.SchemaName,
                            Description = otherAgent.GetDescription() ?? "",
                            HistoryType = "",
                            ConnectionType = "Connected Agent"
                        });
                    }
                }
                else
                {
                    // This agent has no connected agents — it may be a child agent.
                    // List only parent agents that reference this agent via InvokeConnectedAgentTaskAction.
                    foreach (var otherAgent in context.Agents)
                    {
                        if (otherAgent.SchemaName.Equals(agent.SchemaName, StringComparison.OrdinalIgnoreCase))
                            continue;
                        var otherConnected = otherAgent.GetAllConnectedAgentInfos();
                        bool isParent = otherConnected.Any(c =>
                            !string.IsNullOrEmpty(c.BotSchemaName) &&
                            c.BotSchemaName.Equals(agent.SchemaName, StringComparison.OrdinalIgnoreCase));
                        if (isParent)
                        {
                            result.Add(new ConnectedAgentInfo
                            {
                                Name = otherAgent.Name,
                                BotSchemaName = otherAgent.SchemaName,
                                Description = otherAgent.GetDescription() ?? "",
                                HistoryType = "",
                                ConnectionType = "Parent Agent"
                            });
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Returns all topic-to-topic calls across all topics, with names resolved.
        /// </summary>
        public List<TopicCallInfo> GetTopicDataFlowInfo()
        {
            var allCalls = new List<TopicCallInfo>();
            var topicNameLookup = agent.GetTopics().ToDictionary(
                t => t.SchemaName,
                t => t.Name,
                StringComparer.OrdinalIgnoreCase);

            foreach (var topic in agent.GetTopics())
            {
                foreach (var call in topic.GetTopicCalls())
                {
                    call.SourceTopicName = topic.Name;
                    if (topicNameLookup.TryGetValue(call.TargetTopicSchemaName, out string targetName))
                        call.TargetTopicName = targetName;
                    allCalls.Add(call);
                }
            }
            return allCalls;
        }

        /// <summary>
        /// Returns a map of Global variable name → list of topic usages (read/write).
        /// </summary>
        public Dictionary<string, List<VariableUsageEntry>> GetGlobalVariableUsageMap()
        {
            var map = new Dictionary<string, List<VariableUsageEntry>>(StringComparer.OrdinalIgnoreCase);
            foreach (var topic in agent.GetTopics())
            {
                foreach (var varRef in topic.GetVariableReferences())
                {
                    if (!varRef.VariableName.StartsWith("Global.", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (!map.TryGetValue(varRef.VariableName, out var entries))
                    {
                        entries = new List<VariableUsageEntry>();
                        map[varRef.VariableName] = entries;
                    }
                    entries.Add(new VariableUsageEntry
                    {
                        TopicName = topic.Name,
                        AccessType = varRef.AccessType,
                        Context = varRef.Context
                    });
                }
            }
            return map;
        }
    }
}