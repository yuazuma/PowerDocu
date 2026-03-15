using System;
using System.Collections.Generic;
using System.IO;
using PowerDocu.Common;
using PowerDocu.AgentDocumenter;
using PowerDocu.AIModelDocumenter;
using PowerDocu.AppDocumenter;
using PowerDocu.AppModuleDocumenter;
using PowerDocu.FlowDocumenter;

namespace PowerDocu.SolutionDocumenter
{
    /// <summary>
    /// Orchestrates the two-phase documentation pipeline for solution zip files:
    /// Phase 1 - Parse all components (solution, flows, apps, agents, customizations)
    /// Phase 2 - Generate all documentation output with full cross-reference access
    /// </summary>
    public static class SolutionDocumentationGenerator
    {
        public static void GenerateDocumentation(string filePath, bool fullDocumentation, ConfigHelper config, string outputPath = null)
        {
            if (!File.Exists(filePath))
            {
                NotificationHelper.SendNotification($"File not found: {filePath}");
                return;
            }

            DateTime startDocGeneration = DateTime.Now;

            // ── Phase 1: Parse everything ──────────────────────────────────
            NotificationHelper.SendNotification("Phase 1: Parsing all components...");

            DocumentationContext context = new DocumentationContext
            {
                Config = config,
                FullDocumentation = fullDocumentation,
                OutputPath = outputPath
            };

            // Parse flows
            var (flows, flowPath) = FlowDocumentationGenerator.ParseFlows(filePath, outputPath);
            context.Flows = flows ?? new List<FlowEntity>();

            // Parse apps
            var (apps, appPath) = AppDocumentationGenerator.ParseApps(filePath, outputPath);
            context.Apps = apps ?? new List<AppEntity>();

            // Parse agents
            var (agents, agentPath) = AgentDocumentationGenerator.ParseAgents(filePath, outputPath);
            context.Agents = agents ?? new List<AgentEntity>();

            // Parse solution metadata and customizations (provides tables, roles, views, flow names, app names, etc.)
            SolutionParser solutionParser = new SolutionParser(filePath);
            if (solutionParser.solution != null)
            {
                context.Solution = solutionParser.solution;
                context.Customizations = solutionParser.solution.Customizations;
                if (context.Customizations != null)
                {
                    context.Tables = context.Customizations.getEntities();
                    context.Roles = context.Customizations.getRoles();
                    // Extract AppModules from customizations
                    if (config.documentModelDrivenApps)
                    {
                        context.AppModules = context.Customizations.getAppModules() ?? new List<AppModuleEntity>();
                    }

                    // Enrich flows with ModernFlowType from customizations.xml
                    foreach (FlowEntity flow in context.Flows)
                    {
                        if (!string.IsNullOrEmpty(flow.ID))
                        {
                            flow.modernFlowType = context.Customizations.getModernFlowTypeById(flow.ID);
                        }
                    }
                }
            }


            NotificationHelper.SendNotification(
                $"Phase 1 complete: {context.Flows.Count} flow(s), {context.Apps.Count} app(s), " +
                $"{context.Agents.Count} agent(s), {context.AppModules.Count} app module(s), " +
                $"{context.Tables.Count} table(s), {context.Roles.Count} role(s)."
            );

            // ── Phase 2: Generate all documentation ────────────────────────
            NotificationHelper.SendNotification("Phase 2: Generating documentation...");

            // Compute centralised solution base path so that all sub-documenters
            // write into the same Solution folder, regardless of how individual
            // parsers classify the package.
            string solutionBasePath = outputPath == null
                ? Path.GetDirectoryName(filePath) + @"\Solution " + CharsetHelper.GetSafeName(Path.GetFileNameWithoutExtension(filePath))
                : outputPath + @"\" + CharsetHelper.GetSafeName(Path.GetFileNameWithoutExtension(filePath));

            // Generate flow documentation
            if (flows != null)
            {
                FlowDocumentationGenerator.GenerateOutput(context, solutionBasePath);
            }

            // Generate app documentation
            if (apps != null)
            {
                AppDocumentationGenerator.GenerateOutput(context, solutionBasePath);
            }

            // Generate agent documentation
            if (agents != null)
            {
                AgentDocumentationGenerator.GenerateOutput(context, solutionBasePath);
            }

            // Generate AI Model documentation
            if (config.documentSolution && context.Solution != null)
            {
                AIModelDocumentationGenerator.GenerateOutput(context, solutionBasePath);
            }

            // Generate solution-level documentation (solution overview, model-driven apps, Dataverse graph)
            if (config.documentSolution && context.Solution != null)
            {
                string solutionPath = solutionBasePath + @"\";

                // Generate Model-Driven App documentation
                if (config.documentModelDrivenApps)
                {
                    AppModuleDocumentationGenerator.GenerateOutput(context, solutionPath);
                }

                // Generate solution overview documentation
                SolutionDocumentationContent solutionContent = new SolutionDocumentationContent(context, solutionPath);
                DataverseGraphBuilder dataverseGraphBuilder = new DataverseGraphBuilder(solutionContent);

                // Generate solution component relationship graph
                SolutionComponentGraphBuilder componentGraphBuilder = new SolutionComponentGraphBuilder(
                    solutionContent, solutionPath, config.showAllComponentsInGraph);
                componentGraphBuilder.Build();

                if (fullDocumentation)
                {
                    if (config.outputFormat.Equals(OutputFormatHelper.Word) || config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Solution documentation");
                        SolutionWordDocBuilder wordzip = new SolutionWordDocBuilder(solutionContent, config.wordTemplate, config.documentDefaultColumns, config.addTableOfContents);
                    }
                    if (config.outputFormat.Equals(OutputFormatHelper.Markdown) || config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        SolutionMarkdownBuilder mdDoc = new SolutionMarkdownBuilder(solutionContent, config.documentDefaultColumns);
                    }
                    if (config.outputFormat.Equals(OutputFormatHelper.Html) || config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating HTML Solution documentation");
                        SolutionHtmlBuilder htmlDoc = new SolutionHtmlBuilder(solutionContent, config.documentDefaultColumns);
                    }
                    FormSvgBuilder.ClearCache();
                }
            }

            DateTime endDocGeneration = DateTime.Now;
            NotificationHelper.SendNotification($"Documentation completed for {filePath}. Total time: {(endDocGeneration - startDocGeneration).TotalSeconds} seconds.");
        }
    }
}