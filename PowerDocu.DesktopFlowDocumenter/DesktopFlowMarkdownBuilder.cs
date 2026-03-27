using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerDocu.Common;
using Grynwald.MarkdownGenerator;

namespace PowerDocu.DesktopFlowDocumenter
{
    class DesktopFlowMarkdownBuilder : MarkdownBuilder
    {
        private readonly DesktopFlowDocumentationContent content;
        private readonly string mainDocumentFileName;
        private readonly MdDocument mainDocument;
        private readonly DocumentSet<MdDocument> set;

        public DesktopFlowMarkdownBuilder(DesktopFlowDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            mainDocumentFileName = ("desktopflow-" + content.filename + ".md").Replace(" ", "-");
            set = new DocumentSet<MdDocument>();
            mainDocument = set.CreateMdDocument(mainDocumentFileName);

            addOverview();
            addFlowDiagram();
            addActionSteps();
            addVariables();
            if (content.flow.ControlFlowBlocks.Count > 0)
                addControlFlow();
            if (content.flow.Subflows.Count > 0)
                addSubflows();
            addModules();
            if (content.flow.Connectors.Count > 0)
                addConnectors();
            if (content.flow.EnvironmentVariables.Count > 0)
                addEnvironmentVariables();
            addProperties();

            set.Save(content.folderPath);
            NotificationHelper.SendNotification("Created Markdown documentation for Desktop Flow: " + content.flow.GetDisplayName());
        }

        private void addOverview()
        {
            mainDocument.Root.Add(new MdHeading(content.flow.GetDisplayName(), 1));

            if (content.context?.Solution != null)
            {
                if (content.context?.Config?.documentSolution == true)
                    mainDocument.Root.Add(new MdParagraph(new MdCompositeSpan(new MdTextSpan("Solution: "), new MdLinkSpan(content.context.Solution.UniqueName, "../" + CrossDocLinkHelper.GetSolutionDocMdPath(content.context.Solution.UniqueName)))));
                else
                    mainDocument.Root.Add(new MdParagraph(new MdTextSpan("Solution: " + content.context.Solution.UniqueName)));
            }

            List<MdTableRow> tableRows = new List<MdTableRow>
            {
                new MdTableRow("Name", content.flow.GetDisplayName()),
                new MdTableRow("State", content.flow.GetStateLabel()),
                new MdTableRow("Action Steps", content.flow.ActionSteps.Count.ToString()),
                new MdTableRow("Variables", content.flow.Variables.Count.ToString()),
                new MdTableRow("Modules Used", content.flow.Modules.Count.ToString())
            };
            if (content.flow.Subflows.Count > 0)
                tableRows.Add(new MdTableRow("Subflows", content.flow.Subflows.Count.ToString()));
            if (!string.IsNullOrEmpty(content.flow.GetEngineVersionString()))
                tableRows.Add(new MdTableRow("Engine Version", content.flow.GetEngineVersionString()));
            if (content.flow.PowerFxEnabled)
                tableRows.Add(new MdTableRow("Power Fx", "Enabled" + (!string.IsNullOrEmpty(content.flow.PowerFxVersion) ? $" (v{content.flow.PowerFxVersion})" : "")));
            if (!string.IsNullOrEmpty(content.flow.Description))
                tableRows.Add(new MdTableRow("Description", content.flow.Description));
            tableRows.Add(new MdTableRow(content.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion()));

            mainDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
        }

        private void addFlowDiagram()
        {
            string graphFile = "desktopflow-detailed.svg";
            if (File.Exists(Path.Combine(content.folderPath, graphFile)))
            {
                mainDocument.Root.Add(new MdHeading("Flow Diagram", 2));
                mainDocument.Root.Add(new MdParagraph(new MdRawMarkdownSpan($"[![Desktop Flow Diagram]({graphFile})]({graphFile})")));
            }
        }

        private void addActionSteps()
        {
            mainDocument.Root.Add(new MdHeading(content.headerActionSteps, 2));

            if (content.flow.ActionSteps.Count == 0)
            {
                mainDocument.Root.Add(new MdParagraph(new MdTextSpan("No action steps found.")));
                return;
            }

            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This Desktop Flow has {content.flow.ActionSteps.Count} action step(s).")));

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (var step in content.flow.ActionSteps.OrderBy(s => s.Order))
            {
                string parameters = step.Parameters.Count > 0
                    ? string.Join(", ", step.Parameters.Select(p => $"{p.Key}: {p.Value}"))
                    : "";
                string outputs = step.OutputVariables.Count > 0
                    ? string.Join(", ", step.OutputVariables)
                    : "";

                tableRows.Add(new MdTableRow(
                    step.Order.ToString(),
                    step.ModuleName,
                    step.FullActionName,
                    parameters,
                    outputs
                ));
            }
            mainDocument.Root.Add(new MdTable(
                new MdTableRow("#", "Module", "Action", "Parameters", "Output Variables"),
                tableRows));
        }

        private void addVariables()
        {
            mainDocument.Root.Add(new MdHeading(content.headerVariables, 2));

            if (content.flow.Variables.Count == 0)
            {
                mainDocument.Root.Add(new MdParagraph(new MdTextSpan("No variables found.")));
                return;
            }

            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This Desktop Flow has {content.flow.Variables.Count} variable(s).")));

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (var variable in content.flow.Variables.OrderBy(v => v.Name))
            {
                tableRows.Add(new MdTableRow(
                    variable.Name,
                    variable.Type ?? "",
                    variable.IsInput ? "Yes" : "",
                    variable.IsOutput ? "Yes" : "",
                    variable.IsSensitive ? "Yes" : "",
                    variable.InitialValue ?? ""
                ));
            }
            mainDocument.Root.Add(new MdTable(
                new MdTableRow("Name", "Type", "Input", "Output", "Sensitive", "Initial Value"),
                tableRows));
        }

        private void addControlFlow()
        {
            mainDocument.Root.Add(new MdHeading(content.headerControlFlow, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This Desktop Flow has {content.flow.ControlFlowBlocks.Count} control flow block(s).")));

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (var block in content.flow.ControlFlowBlocks)
            {
                tableRows.Add(new MdTableRow(
                    block.Type,
                    block.Condition ?? "",
                    block.StartLine.ToString(),
                    block.NestingLevel.ToString()
                ));
            }
            mainDocument.Root.Add(new MdTable(
                new MdTableRow("Type", "Condition", "Line", "Nesting Level"),
                tableRows));
        }

        private void addSubflows()
        {
            mainDocument.Root.Add(new MdHeading(content.headerSubflows, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This Desktop Flow has {content.flow.Subflows.Count} subflow(s).")));

            foreach (var subflow in content.flow.Subflows)
            {
                string subflowTitle = subflow.Name + (subflow.IsGlobal ? " (Global)" : "");
                mainDocument.Root.Add(new MdHeading(subflowTitle, 3));

                // Subflow flow diagram
                string subflowDiagramFile = "desktopflow-subflow-" + CharsetHelper.GetSafeName(subflow.Name) + "-detailed.svg";
                if (File.Exists(Path.Combine(content.folderPath, subflowDiagramFile)))
                {
                    mainDocument.Root.Add(new MdParagraph(new MdRawMarkdownSpan($"[![Subflow Diagram: {subflow.Name}]({subflowDiagramFile})]({subflowDiagramFile})")));
                }

                // Subflow action steps
                if (subflow.ActionSteps.Count > 0)
                {
                    mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"{subflow.ActionSteps.Count} action step(s).")));
                    List<MdTableRow> stepRows = new List<MdTableRow>();
                    foreach (var step in subflow.ActionSteps.OrderBy(s => s.Order))
                    {
                        string parameters = step.Parameters.Count > 0
                            ? string.Join(", ", step.Parameters.Select(p => $"{p.Key}: {p.Value}"))
                            : "";
                        string outputs = step.OutputVariables.Count > 0
                            ? string.Join(", ", step.OutputVariables)
                            : "";
                        stepRows.Add(new MdTableRow(step.Order.ToString(), step.ModuleName, step.FullActionName, parameters, outputs));
                    }
                    mainDocument.Root.Add(new MdTable(
                        new MdTableRow("#", "Module", "Action", "Parameters", "Output Variables"),
                        stepRows));
                }

                // Subflow variables
                if (subflow.Variables.Count > 0)
                {
                    mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"{subflow.Variables.Count} variable(s).")));
                    List<MdTableRow> varRows = new List<MdTableRow>();
                    foreach (var variable in subflow.Variables.OrderBy(v => v.Name))
                    {
                        varRows.Add(new MdTableRow(
                            variable.Name,
                            variable.Type ?? "",
                            variable.IsInput ? "Yes" : "",
                            variable.IsOutput ? "Yes" : "",
                            variable.IsSensitive ? "Yes" : "",
                            variable.InitialValue ?? ""
                        ));
                    }
                    mainDocument.Root.Add(new MdTable(
                        new MdTableRow("Name", "Type", "Input", "Output", "Sensitive", "Initial Value"),
                        varRows));
                }

                // Subflow control flow
                if (subflow.ControlFlowBlocks.Count > 0)
                {
                    mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"{subflow.ControlFlowBlocks.Count} control flow block(s).")));
                    List<MdTableRow> cfRows = new List<MdTableRow>();
                    foreach (var block in subflow.ControlFlowBlocks)
                    {
                        cfRows.Add(new MdTableRow(
                            block.Type,
                            block.Condition ?? "",
                            block.StartLine.ToString(),
                            block.NestingLevel.ToString()
                        ));
                    }
                    mainDocument.Root.Add(new MdTable(
                        new MdTableRow("Type", "Condition", "Line", "Nesting Level"),
                        cfRows));
                }
            }
        }

        private void addModules()
        {
            mainDocument.Root.Add(new MdHeading(content.headerModules, 2));

            if (content.flow.Modules.Count == 0)
            {
                mainDocument.Root.Add(new MdParagraph(new MdTextSpan("No modules referenced.")));
                return;
            }

            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This Desktop Flow references {content.flow.Modules.Count} module(s).")));

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (var module in content.flow.Modules.OrderBy(m => m.Name))
            {
                int actionCount = content.flow.ActionSteps.Count(a => a.ModuleName == module.Name);
                tableRows.Add(new MdTableRow(
                    module.Name,
                    module.AssemblyName ?? "",
                    module.Version ?? "",
                    actionCount.ToString()
                ));
            }
            mainDocument.Root.Add(new MdTable(
                new MdTableRow("Module", "Assembly", "Version", "Actions Used"),
                tableRows));
        }

        private void addConnectors()
        {
            mainDocument.Root.Add(new MdHeading(content.headerConnectors, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This Desktop Flow uses {content.flow.Connectors.Count} cloud connector(s).")));

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (var connector in content.flow.Connectors)
            {
                tableRows.Add(new MdTableRow(
                    connector.ConnectorId ?? "",
                    connector.Name ?? "",
                    connector.Title ?? ""
                ));
            }
            mainDocument.Root.Add(new MdTable(
                new MdTableRow("Connector ID", "Name", "Title"),
                tableRows));
        }

        private void addEnvironmentVariables()
        {
            mainDocument.Root.Add(new MdHeading(content.headerEnvironmentVariables, 2));

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (var envVar in content.flow.EnvironmentVariables)
            {
                tableRows.Add(new MdTableRow(
                    envVar.Name ?? "",
                    envVar.Type ?? "",
                    envVar.Value ?? ""
                ));
            }
            mainDocument.Root.Add(new MdTable(
                new MdTableRow("Name", "Type", "Value"),
                tableRows));
        }

        private void addProperties()
        {
            mainDocument.Root.Add(new MdHeading(content.headerProperties, 2));

            List<MdTableRow> tableRows = new List<MdTableRow>
            {
                new MdTableRow("ID", content.flow.ID ?? ""),
                new MdTableRow("Schema Version", content.flow.SchemaVersion ?? ""),
                new MdTableRow("UI Flow Type", content.flow.UIFlowType.ToString()),
                new MdTableRow("State", content.flow.GetStateLabel()),
                new MdTableRow("Is Customizable", content.flow.IsCustomizable ? "Yes" : "No")
            };
            if (!string.IsNullOrEmpty(content.flow.IntroducedVersion))
                tableRows.Add(new MdTableRow("Introduced Version", content.flow.IntroducedVersion));
            if (!string.IsNullOrEmpty(content.flow.GetEngineVersionString()))
                tableRows.Add(new MdTableRow("Engine Version", content.flow.GetEngineVersionString()));
            if (content.flow.PowerFxEnabled)
                tableRows.Add(new MdTableRow("Power Fx Version", content.flow.PowerFxVersion ?? ""));

            mainDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
        }

    }
}
