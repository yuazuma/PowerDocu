using System;
using System.IO;
using PowerDocu.Common;

namespace PowerDocu.DesktopFlowDocumenter
{
    public static class DesktopFlowDocumentationGenerator
    {
        public static void GenerateOutput(DocumentationContext context, string path)
        {
            if (context.DesktopFlows == null || context.DesktopFlows.Count == 0 || !context.Config.documentDesktopFlows) return;

            DateTime startDocGeneration = DateTime.Now;
            NotificationHelper.SendNotification($"Found {context.DesktopFlows.Count} Desktop Flow(s) in the solution.");

            if (context.FullDocumentation)
            {
                foreach (DesktopFlowEntity flow in context.DesktopFlows)
                {
                    DesktopFlowDocumentationContent content = new DesktopFlowDocumentationContent(flow, path, context);

                    // Generate flow diagram graph for main flow
                    if (flow.ActionSteps.Count > 0)
                    {
                        GraphBuilder graphBuilder = new GraphBuilder(flow, content.folderPath);
                        graphBuilder.buildDetailedGraph();
                    }

                    // Generate flow diagram graphs for each subflow
                    foreach (var subflow in flow.Subflows)
                    {
                        if (subflow.ActionSteps.Count > 0)
                        {
                            GraphBuilder subflowGraphBuilder = new GraphBuilder(subflow, flow.GetDisplayName(), content.folderPath);
                            subflowGraphBuilder.buildDetailedGraph();
                        }
                    }

                    string wordTemplate = (!String.IsNullOrEmpty(context.Config.wordTemplate) && File.Exists(context.Config.wordTemplate))
                        ? context.Config.wordTemplate : null;
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Word) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Word documentation for Desktop Flow: " + flow.GetDisplayName());
                        DesktopFlowWordDocBuilder wordDoc = new DesktopFlowWordDocBuilder(content, wordTemplate);
                    }
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Markdown) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Markdown documentation for Desktop Flow: " + flow.GetDisplayName());
                        DesktopFlowMarkdownBuilder markdownDoc = new DesktopFlowMarkdownBuilder(content);
                    }
                    if (context.Config.outputFormat.Equals(OutputFormatHelper.Html) || context.Config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating HTML documentation for Desktop Flow: " + flow.GetDisplayName());
                        DesktopFlowHtmlBuilder htmlDoc = new DesktopFlowHtmlBuilder(content);
                    }
                    context.Progress?.Increment("DesktopFlows");
                }
            }
            else
            {
                context.Progress?.Complete("DesktopFlows");
            }

            DateTime endDocGeneration = DateTime.Now;
            NotificationHelper.SendNotification(
                $"DesktopFlowDocumenter: Processed {context.DesktopFlows.Count} Desktop Flow(s) in {(endDocGeneration - startDocGeneration).TotalSeconds} seconds."
            );
        }
    }
}
