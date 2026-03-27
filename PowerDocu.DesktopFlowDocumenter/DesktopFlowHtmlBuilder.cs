using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using PowerDocu.Common;

namespace PowerDocu.DesktopFlowDocumenter
{
    class DesktopFlowHtmlBuilder : HtmlBuilder
    {
        private readonly DesktopFlowDocumentationContent content;
        private readonly string mainFileName;

        public DesktopFlowHtmlBuilder(DesktopFlowDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            WriteDefaultStylesheet(content.folderPath);
            mainFileName = ("desktopflow-" + content.filename + ".html").Replace(" ", "-");

            addOverviewPage();
            NotificationHelper.SendNotification("Created HTML documentation for Desktop Flow: " + content.flow.GetDisplayName());
        }

        private string getNavigationHtml()
        {
            var navItemsList = new List<(string label, string href)>();
            if (content.context?.Solution != null)
            {
                if (content.context?.Config?.documentSolution == true)
                    navItemsList.Add(("Solution", "../" + CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName)));
                else
                    navItemsList.Add((content.context.Solution.UniqueName, ""));
            }
            navItemsList.Add(("Overview", "#overview"));
            if (File.Exists(Path.Combine(content.folderPath, "desktopflow-detailed.svg")))
                navItemsList.Add(("Flow Diagram", "#flow-diagram"));
            navItemsList.Add(("Action Steps", "#action-steps"));
            navItemsList.Add(("Variables", "#variables"));
            if (content.flow.ControlFlowBlocks.Count > 0)
                navItemsList.Add(("Control Flow", "#control-flow"));
            if (content.flow.Subflows.Count > 0)
            {
                navItemsList.Add(("Subflows", "#subflows"));
                foreach (var subflow in content.flow.Subflows)
                    navItemsList.Add(("  " + subflow.Name, "#subflow-" + subflow.Name.ToLowerInvariant()));
            }
            navItemsList.Add(("Modules", "#modules"));
            if (content.flow.Connectors.Count > 0)
                navItemsList.Add(("Connectors", "#connectors"));
            if (content.flow.EnvironmentVariables.Count > 0)
                navItemsList.Add(("Environment Variables", "#environment-variables"));
            navItemsList.Add(("Properties", "#properties"));

            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"<div class=\"nav-title\">{Encode(content.flow.GetDisplayName())}</div>");
            sb.Append(NavigationList(navItemsList));
            return sb.ToString();
        }

        private void addOverviewPage()
        {
            StringBuilder body = new StringBuilder();

            // Overview
            body.AppendLine(HeadingWithId(1, content.flow.GetDisplayName(), "overview"));

            body.Append(TableStart("Property", "Value"));
            body.Append(TableRow("Name", content.flow.GetDisplayName()));
            body.Append(TableRow("State", content.flow.GetStateLabel()));
            body.Append(TableRow("Action Steps", content.flow.ActionSteps.Count.ToString()));
            body.Append(TableRow("Variables", content.flow.Variables.Count.ToString()));
            body.Append(TableRow("Modules Used", content.flow.Modules.Count.ToString()));
            if (content.flow.Subflows.Count > 0)
                body.Append(TableRow("Subflows", content.flow.Subflows.Count.ToString()));
            if (!string.IsNullOrEmpty(content.flow.GetEngineVersionString()))
                body.Append(TableRow("Engine Version", content.flow.GetEngineVersionString()));
            if (content.flow.PowerFxEnabled)
                body.Append(TableRow("Power Fx", "Enabled" + (!string.IsNullOrEmpty(content.flow.PowerFxVersion) ? $" (v{content.flow.PowerFxVersion})" : "")));
            if (!string.IsNullOrEmpty(content.flow.Description))
                body.Append(TableRow("Description", content.flow.Description));
            body.Append(TableRow(content.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion()));
            body.AppendLine(TableEnd());

            // Flow Diagram
            string flowDiagramFile = "desktopflow-detailed.svg";
            if (File.Exists(Path.Combine(content.folderPath, flowDiagramFile)))
            {
                body.AppendLine(HeadingWithId(2, "Flow Diagram", "flow-diagram"));
                body.AppendLine(ParagraphRaw($"<a href=\"{Encode(flowDiagramFile)}\" target=\"_blank\">{Image("Desktop Flow Diagram", flowDiagramFile)}</a>"));
            }

            // Action Steps
            addActionSteps(body);

            // Variables
            addVariables(body);

            // Control Flow
            if (content.flow.ControlFlowBlocks.Count > 0)
                addControlFlow(body);

            // Subflows
            if (content.flow.Subflows.Count > 0)
                addSubflows(body);

            // Modules
            addModules(body);

            // Connectors
            if (content.flow.Connectors.Count > 0)
                addConnectors(body);

            // Environment Variables
            if (content.flow.EnvironmentVariables.Count > 0)
                addEnvironmentVariables(body);

            // Properties
            addProperties(body);

            SaveHtmlFile(Path.Combine(content.folderPath, mainFileName),
                WrapInHtmlPage($"Desktop Flow - {content.flow.GetDisplayName()}", body.ToString(), getNavigationHtml()));
        }

        private void addActionSteps(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, content.headerActionSteps, "action-steps"));

            if (content.flow.ActionSteps.Count == 0)
            {
                body.AppendLine("<p>No action steps found.</p>");
                return;
            }

            body.AppendLine($"<p>This Desktop Flow has {content.flow.ActionSteps.Count} action step(s).</p>");

            body.Append(TableStart("#", "Module", "Action", "Parameters", "Output Variables"));
            foreach (var step in content.flow.ActionSteps.OrderBy(s => s.Order))
            {
                string indent = new string('\u00A0', step.NestingLevel * 4);
                string parameters = step.Parameters.Count > 0
                    ? string.Join(", ", step.Parameters.Select(p => $"{p.Key}: {p.Value}"))
                    : "";
                string outputs = step.OutputVariables.Count > 0
                    ? string.Join(", ", step.OutputVariables)
                    : "";

                body.Append(TableRow(
                    step.Order.ToString(),
                    step.ModuleName,
                    indent + step.FullActionName,
                    parameters,
                    outputs
                ));
            }
            body.AppendLine(TableEnd());
        }

        private void addVariables(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, content.headerVariables, "variables"));

            if (content.flow.Variables.Count == 0)
            {
                body.AppendLine("<p>No variables found.</p>");
                return;
            }

            body.AppendLine($"<p>This Desktop Flow has {content.flow.Variables.Count} variable(s).</p>");

            body.Append(TableStart("Name", "Type", "Input", "Output", "Sensitive", "Initial Value"));
            foreach (var variable in content.flow.Variables.OrderBy(v => v.Name))
            {
                body.Append(TableRow(
                    variable.Name,
                    variable.Type ?? "",
                    variable.IsInput ? "Yes" : "",
                    variable.IsOutput ? "Yes" : "",
                    variable.IsSensitive ? "Yes" : "",
                    variable.InitialValue ?? ""
                ));
            }
            body.AppendLine(TableEnd());
        }

        private void addControlFlow(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, content.headerControlFlow, "control-flow"));
            body.AppendLine($"<p>This Desktop Flow has {content.flow.ControlFlowBlocks.Count} control flow block(s).</p>");

            body.Append(TableStart("Type", "Condition", "Line", "Nesting Level"));
            foreach (var block in content.flow.ControlFlowBlocks)
            {
                body.Append(TableRow(
                    block.Type,
                    block.Condition ?? "",
                    block.StartLine.ToString(),
                    block.NestingLevel.ToString()
                ));
            }
            body.AppendLine(TableEnd());
        }

        private void addSubflows(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, content.headerSubflows, "subflows"));
            body.AppendLine($"<p>This Desktop Flow has {content.flow.Subflows.Count} subflow(s).</p>");

            foreach (var subflow in content.flow.Subflows)
            {
                string subflowTitle = subflow.Name + (subflow.IsGlobal ? " (Global)" : "");
                body.AppendLine(HeadingWithId(3, subflowTitle, "subflow-" + subflow.Name.ToLowerInvariant()));

                // Subflow flow diagram
                string subflowDiagramFile = "desktopflow-subflow-" + CharsetHelper.GetSafeName(subflow.Name) + "-detailed.svg";
                if (File.Exists(Path.Combine(content.folderPath, subflowDiagramFile)))
                {
                    body.AppendLine(ParagraphRaw($"<a href=\"{Encode(subflowDiagramFile)}\" target=\"_blank\">{Image("Subflow Diagram: " + subflow.Name, subflowDiagramFile)}</a>"));
                }

                // Subflow action steps
                if (subflow.ActionSteps.Count > 0)
                {
                    body.AppendLine($"<p>{subflow.ActionSteps.Count} action step(s).</p>");
                    body.Append(TableStart("#", "Module", "Action", "Parameters", "Output Variables"));
                    foreach (var step in subflow.ActionSteps.OrderBy(s => s.Order))
                    {
                        string indent = new string('\u00A0', step.NestingLevel * 4);
                        string parameters = step.Parameters.Count > 0
                            ? string.Join(", ", step.Parameters.Select(p => $"{p.Key}: {p.Value}"))
                            : "";
                        string outputs = step.OutputVariables.Count > 0
                            ? string.Join(", ", step.OutputVariables)
                            : "";

                        body.Append(TableRow(
                            step.Order.ToString(),
                            step.ModuleName,
                            indent + step.FullActionName,
                            parameters,
                            outputs
                        ));
                    }
                    body.AppendLine(TableEnd());
                }

                // Subflow variables
                if (subflow.Variables.Count > 0)
                {
                    body.AppendLine($"<p>{subflow.Variables.Count} variable(s).</p>");
                    body.Append(TableStart("Name", "Type", "Input", "Output", "Sensitive", "Initial Value"));
                    foreach (var variable in subflow.Variables.OrderBy(v => v.Name))
                    {
                        body.Append(TableRow(
                            variable.Name,
                            variable.Type ?? "",
                            variable.IsInput ? "Yes" : "",
                            variable.IsOutput ? "Yes" : "",
                            variable.IsSensitive ? "Yes" : "",
                            variable.InitialValue ?? ""
                        ));
                    }
                    body.AppendLine(TableEnd());
                }

                // Subflow control flow
                if (subflow.ControlFlowBlocks.Count > 0)
                {
                    body.AppendLine($"<p>{subflow.ControlFlowBlocks.Count} control flow block(s).</p>");
                    body.Append(TableStart("Type", "Condition", "Line", "Nesting Level"));
                    foreach (var block in subflow.ControlFlowBlocks)
                    {
                        body.Append(TableRow(
                            block.Type,
                            block.Condition ?? "",
                            block.StartLine.ToString(),
                            block.NestingLevel.ToString()
                        ));
                    }
                    body.AppendLine(TableEnd());
                }
            }
        }

        private void addModules(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, content.headerModules, "modules"));

            if (content.flow.Modules.Count == 0)
            {
                body.AppendLine("<p>No modules referenced.</p>");
                return;
            }

            body.AppendLine($"<p>This Desktop Flow references {content.flow.Modules.Count} module(s).</p>");

            body.Append(TableStart("Module", "Assembly", "Version", "Actions Used"));
            foreach (var module in content.flow.Modules.OrderBy(m => m.Name))
            {
                int actionCount = content.flow.ActionSteps.Count(a => a.ModuleName == module.Name);
                body.Append(TableRow(
                    module.Name,
                    module.AssemblyName ?? "",
                    module.Version ?? "",
                    actionCount.ToString()
                ));
            }
            body.AppendLine(TableEnd());
        }

        private void addConnectors(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, content.headerConnectors, "connectors"));
            body.AppendLine($"<p>This Desktop Flow uses {content.flow.Connectors.Count} cloud connector(s).</p>");

            body.Append(TableStart("Connector ID", "Name", "Title"));
            foreach (var connector in content.flow.Connectors)
            {
                body.Append(TableRow(
                    connector.ConnectorId ?? "",
                    connector.Name ?? "",
                    connector.Title ?? ""
                ));
            }
            body.AppendLine(TableEnd());
        }

        private void addEnvironmentVariables(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, content.headerEnvironmentVariables, "environment-variables"));

            body.Append(TableStart("Name", "Type", "Value"));
            foreach (var envVar in content.flow.EnvironmentVariables)
            {
                body.Append(TableRow(
                    envVar.Name ?? "",
                    envVar.Type ?? "",
                    envVar.Value ?? ""
                ));
            }
            body.AppendLine(TableEnd());
        }

        private void addProperties(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, content.headerProperties, "properties"));
            body.Append(TableStart("Property", "Value"));
            body.Append(TableRow("ID", content.flow.ID ?? ""));
            body.Append(TableRow("Schema Version", content.flow.SchemaVersion ?? ""));
            body.Append(TableRow("UI Flow Type", content.flow.UIFlowType.ToString()));
            body.Append(TableRow("State", content.flow.GetStateLabel()));
            body.Append(TableRow("Is Customizable", content.flow.IsCustomizable ? "Yes" : "No"));
            if (!string.IsNullOrEmpty(content.flow.IntroducedVersion))
                body.Append(TableRow("Introduced Version", content.flow.IntroducedVersion));
            if (!string.IsNullOrEmpty(content.flow.GetEngineVersionString()))
                body.Append(TableRow("Engine Version", content.flow.GetEngineVersionString()));
            if (content.flow.PowerFxEnabled)
                body.Append(TableRow("Power Fx Version", content.flow.PowerFxVersion ?? ""));
            body.AppendLine(TableEnd());
        }

    }
}
