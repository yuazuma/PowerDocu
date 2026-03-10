using System;
using System.Collections.Generic;
using System.IO;
using PowerDocu.Common;

namespace PowerDocu.AgentDocumenter
{
    public static class AgentDocumentationGenerator
    {
        /// <summary>
        /// Parses agents from the given file without generating documentation output.
        /// Returns the parsed agents and the resolved output path.
        /// </summary>
        public static (List<AgentEntity> Agents, string Path) ParseAgents(string filePath, string outputPath = null)
        {
            if (!File.Exists(filePath))
            {
                NotificationHelper.SendNotification("File not found: " + filePath);
                return (null, null);
            }

            string path =
                outputPath == null
                    ? Path.GetDirectoryName(filePath)
                    : $"{outputPath}/{Path.GetFileNameWithoutExtension(filePath)}";
            path += @"\Solution " + CharsetHelper.GetSafeName(Path.GetFileNameWithoutExtension(filePath));
            AgentParser agentParserFromZip = new AgentParser(filePath);
            List<AgentEntity> agents = agentParserFromZip.getAgents();
            NotificationHelper.SendNotification($"AgentParser: Parsed {agents.Count} agent(s) from {filePath}.");
            return (agents, path);
        }

        /// <summary>
        /// Generates documentation output for pre-parsed agents using the DocumentationContext.
        /// </summary>
        public static void GenerateOutput(DocumentationContext context, string path)
        {
            if (context.Agents == null || !context.Config.documentAgents)
            {
                if (!context.Config.documentAgents)
                    NotificationHelper.SendNotification("Agent documentation is disabled in configuration.");
                return;
            }

            DateTime startDocGeneration = DateTime.Now;
            foreach (AgentEntity agent in context.Agents)
            {
                string folderPath =
                    path + CharsetHelper.GetSafeName(@"\AgentDoc " + agent.Name + @"\");
                Directory.CreateDirectory(folderPath);
                string topicsFolderPath = folderPath + @"Topics\";
                Directory.CreateDirectory(topicsFolderPath);
                foreach (BotComponent topic in agent.GetTopics())
                {
                    GraphBuilder graphBuilder = new GraphBuilder(agent.Name, topic, topicsFolderPath, context);
                    graphBuilder.buildDetailedGraph();
                }

                if (context.FullDocumentation)
                {
                    AgentDocumentationContent content = new AgentDocumentationContent(agent, path, context);
                    string wordTemplate = (!String.IsNullOrEmpty(context.Config.wordTemplate) && File.Exists(context.Config.wordTemplate))
                        ? context.Config.wordTemplate : null;
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Word) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Word documentation");
                        AgentWordDocBuilder wordzip = new AgentWordDocBuilder(content, wordTemplate);
                    }
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Markdown) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Markdown documentation");
                        AgentMarkdownBuilder agentMarkdownBuilder = new AgentMarkdownBuilder(content);
                    }
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Html) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating HTML documentation");
                        AgentHtmlBuilder agentHtmlBuilder = new AgentHtmlBuilder(content);
                    }
                }
            }
            DateTime endDocGeneration = DateTime.Now;
            NotificationHelper.SendNotification(
                $"AgentDocumenter: Generated documentation for {context.Agents.Count} agent(s) in {(endDocGeneration - startDocGeneration).TotalSeconds} seconds."
            );
        }

        /// <summary>
        /// Legacy method: parses and generates documentation in one step.
        /// </summary>
        public static List<AgentEntity> GenerateDocumentation(
            string filePath,
            bool fullDocumentation,
            ConfigHelper config,
            string outputPath = null
        )
        {
            var (agents, path) = ParseAgents(filePath, outputPath);
            if (agents == null) return null;

            var context = new DocumentationContext
            {
                Agents = agents,
                Config = config,
                FullDocumentation = fullDocumentation
            };
            GenerateOutput(context, path);
            return agents;
        }
    }
}
