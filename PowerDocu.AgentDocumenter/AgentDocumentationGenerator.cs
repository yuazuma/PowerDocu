using System;
using System.Collections.Generic;
using System.IO;
using PowerDocu.Common;

namespace PowerDocu.AgentDocumenter
{
    public static class AgentDocumentationGenerator
    {
        public static List<AgentEntity> GenerateDocumentation(
            string filePath,
            bool fullDocumentation,
            ConfigHelper config,
            string outputPath = null
        )
        {
            if (File.Exists(filePath))
            {
                string path =
                    outputPath == null
                        ? Path.GetDirectoryName(filePath)
                        : $"{outputPath}/{Path.GetFileNameWithoutExtension(filePath)}";
                path += @"\Solution " + CharsetHelper.GetSafeName(Path.GetFileNameWithoutExtension(filePath));
                DateTime startDocGeneration = DateTime.Now;
                AgentParser agentParserFromZip = new AgentParser(filePath);

                List<AgentEntity> agents = agentParserFromZip.getAgents();

                if (!config.documentAgents)
                {
                    NotificationHelper.SendNotification("Agent documentation is disabled in configuration.");
                    return agents;
                }

                foreach (AgentEntity agent in agents)
                {
                    string folderPath =
                        path + CharsetHelper.GetSafeName(@"\AgentDoc " + agent.Name + @"\");
                    Directory.CreateDirectory(folderPath);
                    //create topic diagrams in Topics subfolder
                    string topicsFolderPath = folderPath + @"Topics\";
                    Directory.CreateDirectory(topicsFolderPath);
                    foreach (BotComponent topic in agent.GetTopics())
                    {
                        GraphBuilder graphBuilder = new GraphBuilder(agent.Name, topic, topicsFolderPath);
                        graphBuilder.buildDetailedGraph();
                    }

                    if (fullDocumentation)
                    {
                        AgentDocumentationContent content = new AgentDocumentationContent(agent, path);
                        if (config.outputFormat.Equals(OutputFormatHelper.Word) || config.outputFormat.Equals(OutputFormatHelper.All))
                        {
                            //create the Word document
                            NotificationHelper.SendNotification("Creating Word documentation");
                            string wordTemplate = null;
                            if (!String.IsNullOrEmpty(config.wordTemplate) && File.Exists(config.wordTemplate))
                            {
                                wordTemplate = config.wordTemplate;
                            }
                            AgentWordDocBuilder wordzip = new AgentWordDocBuilder(content, wordTemplate);
                        }
                        if (config.outputFormat.Equals(OutputFormatHelper.Markdown) || config.outputFormat.Equals(OutputFormatHelper.All))
                        {
                            NotificationHelper.SendNotification("Creating Markdown documentation");
                            AgentMarkdownBuilder agentMarkdownBuilder = new AgentMarkdownBuilder(content);
                        }
                        if (config.outputFormat.Equals(OutputFormatHelper.Html) || config.outputFormat.Equals(OutputFormatHelper.All))
                        {
                            NotificationHelper.SendNotification("Creating HTML documentation");
                            AgentHtmlBuilder agentHtmlBuilder = new AgentHtmlBuilder(content);
                        }
                    }
                }
                DateTime endDocGeneration = DateTime.Now;
                NotificationHelper.SendNotification(
                    "AgentDocumenter: Created documentation for "
                        + filePath
                        + ". A total of "
                        + agents.Count
                        + " agents were processed in "
                        + (endDocGeneration - startDocGeneration).TotalSeconds
                        + " seconds."
                );
                return agents;
            }
            else
            {
                NotificationHelper.SendNotification("File not found: " + filePath);
            }
            return null;
        }
    }
}
