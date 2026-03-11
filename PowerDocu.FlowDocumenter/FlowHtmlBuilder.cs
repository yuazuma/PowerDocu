using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using PowerDocu.Common;

namespace PowerDocu.FlowDocumenter
{
    class FlowHtmlBuilder : HtmlBuilder
    {
        private readonly string mainFileName, connectionsFileName, variablesFileName, triggerActionsFileName;
        private readonly FlowDocumentationContent content;

        public FlowHtmlBuilder(FlowDocumentationContent contentdocumentation)
        {
            content = contentdocumentation;
            Directory.CreateDirectory(content.folderPath);
            WriteDefaultStylesheet(content.folderPath);

            mainFileName = ("index-" + content.filename + ".html").Replace(" ", "-");
            connectionsFileName = ("connections-" + content.filename + ".html").Replace(" ", "-");
            variablesFileName = ("variables-" + content.filename + ".html").Replace(" ", "-");
            triggerActionsFileName = ("triggersactions-" + content.filename + ".html").Replace(" ", "-");

            addFlowMetadataAndOverview();
            addConnectionReferenceInfo();
            addTriggerInfo();
            addVariablesInfo();
            addActionInfo();
            addFlowDetails();
            NotificationHelper.SendNotification("Created HTML documentation for " + content.metadata.Name);
        }

        private string getNavigationHtml(bool fromSubfolder = false)
        {
            string prefix = fromSubfolder ? "../" : "";
            var navItems = new List<(string label, string href)>();
            if (content.context?.Solution != null)
            {
                string solutionPrefix = fromSubfolder ? "../../" : "../";
                if (content.context?.Config?.documentSolution == true)
                    navItems.Add(("Solution", solutionPrefix + CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName)));
                else
                    navItems.Add((content.context.Solution.UniqueName, ""));
            }
            navItems.AddRange(new (string label, string href)[]
            {
                ("Overview", prefix + mainFileName),
                ("Connection References", prefix + connectionsFileName),
                ("Variables", prefix + variablesFileName),
                ("Triggers & Actions", prefix + triggerActionsFileName)
            });
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"<div class=\"nav-title\">{Encode(content.metadata.Name)}</div>");
            sb.Append(NavigationList(navItems));
            return sb.ToString();
        }

        private string buildMetadataTable()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(TableStart("Flow Name", content.metadata.Name));
            foreach (KeyValuePair<string, string> kvp in content.metadata.metadataTable)
            {
                sb.Append(TableRow(kvp.Key, kvp.Value));
            }
            sb.Append(TableEnd());
            return sb.ToString();
        }

        private void addFlowMetadataAndOverview()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.metadata.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.overview.header));
            body.AppendLine(Paragraph(content.overview.infoText));
            body.AppendLine(ParagraphRaw(Image("Flow Overview Diagram", content.overview.svgFile)));
            SaveHtmlFile(Path.Combine(content.folderPath, mainFileName),
                WrapInHtmlPage(content.metadata.header, body.ToString(), getNavigationHtml()));
        }

        private void addConnectionReferenceInfo()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.metadata.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.connectionReferences.header));
            body.AppendLine(Paragraph(content.connectionReferences.infoText));

            foreach (KeyValuePair<string, Dictionary<string, string>> kvp in content.connectionReferences.connectionTable)
            {
                string connectorUniqueName = kvp.Key;
                ConnectorIcon connectorIcon = ConnectorHelper.getConnectorIcon(connectorUniqueName);
                string displayName = (connectorIcon != null) ? connectorIcon.Name : connectorUniqueName;
                body.AppendLine(Heading(3, displayName));

                string connectorNameHtml = getConnectorNameAndIconHtml(connectorUniqueName, "https://docs.microsoft.com/connectors/" + connectorUniqueName);
                body.Append(TableStart("Connector", ""));
                body.Append(TableRowRaw("Connector", connectorNameHtml));
                foreach (KeyValuePair<string, string> kvp2 in kvp.Value)
                {
                    body.Append(TableRow(kvp2.Key, kvp2.Value));
                }
                body.AppendLine(TableEnd());
            }

            SaveHtmlFile(Path.Combine(content.folderPath, connectionsFileName),
                WrapInHtmlPage("Connections - " + content.metadata.Name, body.ToString(), getNavigationHtml()));
        }

        private string getConnectorNameAndIconHtml(string connectorUniqueName, string url, bool fromSubfolder = false)
        {
            ConnectorIcon connectorIcon = ConnectorHelper.getConnectorIcon(connectorUniqueName);
            string displayName = (connectorIcon != null) ? connectorIcon.Name : connectorUniqueName;
            if (ConnectorHelper.getConnectorIconFile(connectorUniqueName) != "")
            {
                string iconSrc = (fromSubfolder ? "../" : "") + connectorUniqueName + "32.png";
                return $"<a href=\"{Encode(url)}\">{ImageWithClass(connectorUniqueName, iconSrc, "icon-inline")} {Encode(displayName)}</a>";
            }
            return Link(displayName, url);
        }

        private void addVariablesInfo()
        {
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.metadata.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.variables.header));

            foreach (KeyValuePair<string, Dictionary<string, string>> kvp in content.variables.variablesTable)
            {
                body.AppendLine(Heading(3, kvp.Key));
                body.Append(TableStart("Property", "Value"));
                foreach (KeyValuePair<string, string> kvp2 in kvp.Value)
                {
                    if (!kvp2.Key.Equals("Initial Value"))
                    {
                        body.Append(TableRow(kvp2.Key, kvp2.Value));
                    }
                }
                body.AppendLine(TableEnd());

                if (kvp.Value.ContainsKey("Initial Value"))
                {
                    content.variables.initialValTable.TryGetValue(kvp.Key, out Dictionary<string, string> initialValues);
                    if (initialValues?.Count > 0)
                    {
                        body.Append(TableStart("Variable Property", "Initial Value"));
                        foreach (KeyValuePair<string, string> initialVal in initialValues)
                        {
                            body.Append(TableRow(initialVal.Key, initialVal.Value));
                        }
                        body.AppendLine(TableEnd());
                    }
                }

                content.variables.referencesTable.TryGetValue(kvp.Key, out List<ActionNode> references);
                if (references?.Count > 0)
                {
                    body.Append(TableStart("Variable Used In"));
                    foreach (ActionNode action in references.OrderBy(o => o.Name).ToList())
                    {
                        string actionFileName = ("actions/" + CharsetHelper.GetSafeName(action.Name) + ".html").Replace(" ", "-");
                        body.Append(TableRowRaw(Link(action.Name, actionFileName)));
                    }
                    body.AppendLine(TableEnd());
                }
            }

            SaveHtmlFile(Path.Combine(content.folderPath, variablesFileName),
                WrapInHtmlPage("Variables - " + content.metadata.Name, body.ToString(), getNavigationHtml()));
        }

        private void addTriggerInfo()
        {
            // Build a dedicated trigger page
            string triggerDocFileName = ("actions/" + content.trigger.header + ".html").Replace(" ", "-");
            StringBuilder triggerBody = new StringBuilder();
            triggerBody.AppendLine(Heading(1, content.metadata.header));
            triggerBody.AppendLine(buildMetadataTable());
            triggerBody.AppendLine(Heading(2, content.trigger.header));

            triggerBody.Append(TableStart("Property", "Value"));
            foreach (KeyValuePair<string, string> kvp in content.trigger.triggerTable)
            {
                if (kvp.Value == "mergedrow")
                {
                    triggerBody.Append(TableRowRaw($"<td colspan=\"2\"><b>{Encode(kvp.Key)}</b></td>"));
                }
                else
                {
                    triggerBody.Append(TableRow(kvp.Key, kvp.Value));
                }
            }
            triggerBody.AppendLine(TableEnd());

            if (content.trigger.inputs?.Count > 0)
            {
                triggerBody.AppendLine(Heading(3, content.trigger.inputsHeader));
                triggerBody.AppendLine(AddExpressionDetails(content.trigger.inputs));
            }
            if (content.trigger.triggerProperties?.Count > 0)
            {
                triggerBody.AppendLine(Heading(3, "Other Trigger Properties"));
                triggerBody.AppendLine(AddExpressionDetails(content.trigger.triggerProperties));
            }

            Directory.CreateDirectory(Path.Combine(content.folderPath, "actions"));
            WriteDefaultStylesheet(Path.Combine(content.folderPath, "actions"));
            SaveHtmlFile(Path.Combine(content.folderPath, triggerDocFileName),
                WrapInHtmlPage("Trigger - " + content.metadata.Name, triggerBody.ToString(), getNavigationHtml(true)));
        }

        private void addActionInfo()
        {
            StringBuilder mainBody = new StringBuilder();
            mainBody.AppendLine(Heading(2, content.trigger.header));
            string triggerDocFileName = ("actions/" + content.trigger.header + ".html").Replace(" ", "-");
            mainBody.AppendLine(BulletListStart());
            mainBody.AppendLine(BulletItemRaw(Link(content.trigger.header, triggerDocFileName)));
            mainBody.AppendLine(BulletListEnd());

            mainBody.AppendLine(Heading(2, content.actions.header));
            mainBody.AppendLine(Paragraph(content.actions.infoText));
            mainBody.AppendLine(BulletListStart());

            List<ActionNode> actionNodesList = content.actions.actionNodesList;
            Directory.CreateDirectory(Path.Combine(content.folderPath, "actions"));

            foreach (ActionNode action in actionNodesList)
            {
                string actionDocFileName = ("actions/" + CharsetHelper.GetSafeName(action.Name) + ".html").Replace(" ", "-");
                mainBody.AppendLine(BulletItemRaw(Link(action.Name, actionDocFileName)));

                // Build dedicated action page
                StringBuilder actionBody = new StringBuilder();
                actionBody.AppendLine(Heading(1, content.metadata.header));
                actionBody.AppendLine(buildMetadataTable());
                actionBody.AppendLine(Heading(2, action.Name));

                actionBody.Append(TableStart("Property", "Value"));
                actionBody.Append(TableRow("Name", action.Name));
                actionBody.Append(TableRow("Type", action.Type));
                if (!String.IsNullOrEmpty(action.Description))
                    actionBody.Append(TableRow("Description / Note", action.Description));
                if (!String.IsNullOrEmpty(action.Connection))
                    actionBody.Append(TableRowRaw(Encode("Connection"), getConnectorNameAndIconHtml(action.Connection, "https://docs.microsoft.com/connectors/" + action.Connection, true)));
                if (action.actionExpression != null || !String.IsNullOrEmpty(action.Expression))
                    actionBody.Append(TableRowRaw(Encode("Expression"), (action.actionExpression != null) ? AddExpressionTable(action.actionExpression).ToString() : CodeBlock(action.Expression)));
                actionBody.AppendLine(TableEnd());

                // Inputs
                if (action.actionInputs.Count > 0 || !String.IsNullOrEmpty(action.Inputs))
                {
                    actionBody.AppendLine(Heading(3, "Inputs"));
                    actionBody.Append(TableStart("Property", "Value"));
                    if (action.actionInputs.Count > 0)
                    {
                        foreach (Expression actionInput in action.actionInputs)
                        {
                            StringBuilder operandsCell = new StringBuilder();
                            if (actionInput.expressionOperands.Count > 1)
                            {
                                operandsCell.Append("<table class=\"expression-table\">");
                                foreach (object actionInputOperand in actionInput.expressionOperands)
                                {
                                    if (actionInputOperand.GetType() == typeof(Expression))
                                        operandsCell.Append(AddExpressionTable((Expression)actionInputOperand, false));
                                    else
                                        operandsCell.Append("<tr><td>").Append(CodeBlock(actionInputOperand.ToString())).Append("</td></tr>");
                                }
                                operandsCell.Append("</table>");
                            }
                            else
                            {
                                if (actionInput.expressionOperands.Count == 0)
                                    operandsCell.Append("");
                                else
                                {
                                    if (actionInput.expressionOperands[0]?.GetType() == typeof(Expression))
                                        operandsCell.Append(AddExpressionTable((Expression)actionInput.expressionOperands[0]));
                                    else if (actionInput.expressionOperands[0]?.GetType() == typeof(List<object>))
                                    {
                                        operandsCell.Append("<table class=\"expression-table\">");
                                        foreach (object obj in (List<object>)actionInput.expressionOperands[0])
                                        {
                                            if (obj.GetType().Equals(typeof(Expression)))
                                                operandsCell.Append(AddExpressionTable((Expression)obj, false));
                                            else if (obj.GetType().Equals(typeof(List<object>)))
                                            {
                                                foreach (object o in (List<object>)obj)
                                                    operandsCell.Append(AddExpressionTable((Expression)o, false));
                                            }
                                        }
                                        operandsCell.Append("</table>");
                                    }
                                    else
                                        operandsCell.Append(CodeBlock(actionInput.expressionOperands[0]?.ToString()));
                                }
                            }
                            actionBody.Append(TableRowRaw(Encode(actionInput.expressionOperator), operandsCell.ToString()));
                        }
                    }
                    if (!String.IsNullOrEmpty(action.Inputs))
                        actionBody.Append(TableRow("Value", action.Inputs));
                    actionBody.AppendLine(TableEnd());
                }

                // Subactions / Switch
                if (action.Subactions.Count > 0 || action.Elseactions.Count > 0)
                {
                    if (action.Subactions.Count > 0)
                    {
                        actionBody.AppendLine(Heading(3, action.Type == "Switch" ? "Switch Actions" : "Subactions"));
                        if (action.Type == "Switch")
                        {
                            actionBody.Append(TableStart("Case Values", "Action"));
                            foreach (ActionNode subaction in action.Subactions)
                            {
                                if (action.switchRelationship.TryGetValue(subaction, out string switchValue))
                                {
                                    string subActionLink = ("actions/" + CharsetHelper.GetSafeName(subaction.Name) + ".html").Replace(" ", "-");
                                    actionBody.Append(TableRowRaw(Encode(switchValue), Link(subaction.Name, getLinkFromAction(subaction.Name, true))));
                                }
                            }
                            actionBody.AppendLine(TableEnd());
                        }
                        else
                        {
                            actionBody.Append(TableStart("Action"));
                            foreach (ActionNode subaction in action.Subactions)
                            {
                                actionBody.Append(TableRowRaw(Link(subaction.Name, getLinkFromAction(subaction.Name, true))));
                            }
                            actionBody.AppendLine(TableEnd());
                        }
                    }
                    if (action.Elseactions.Count > 0)
                    {
                        actionBody.AppendLine(Heading(3, "Elseactions"));
                        actionBody.Append(TableStart("Elseactions"));
                        foreach (ActionNode elseaction in action.Elseactions)
                        {
                            actionBody.Append(TableRowRaw(Link(elseaction.Name, getLinkFromAction(elseaction.Name, true))));
                        }
                        actionBody.AppendLine(TableEnd());
                    }
                }

                // Next Action(s) Conditions
                if (action.Neighbours.Count > 0)
                {
                    actionBody.AppendLine(Heading(3, "Next Action(s) Conditions"));
                    actionBody.Append(TableStart("Next Action"));
                    foreach (ActionNode nextAction in action.Neighbours)
                    {
                        string[] raConditions = action.nodeRunAfterConditions[nextAction];
                        actionBody.Append(TableRowRaw(Link(nextAction.Name + " [" + string.Join(", ", raConditions) + "]", getLinkFromAction(nextAction.Name, true))));
                    }
                    actionBody.AppendLine(TableEnd());
                }

                SaveHtmlFile(Path.Combine(content.folderPath, actionDocFileName),
                    WrapInHtmlPage("Action: " + action.Name, actionBody.ToString(), getNavigationHtml(true)));
            }

            mainBody.AppendLine(BulletListEnd());

            // Now append to the triggers/actions overview file
            StringBuilder fullBody = new StringBuilder();
            fullBody.AppendLine(Heading(1, content.metadata.header));
            fullBody.AppendLine(buildMetadataTable());
            fullBody.Append(mainBody);
            SaveHtmlFile(Path.Combine(content.folderPath, triggerActionsFileName),
                WrapInHtmlPage("Triggers & Actions - " + content.metadata.Name, fullBody.ToString(), getNavigationHtml()));
        }

        private void addFlowDetails()
        {
            // Append detailed flow diagram to the main overview page
            StringBuilder body = new StringBuilder();
            body.AppendLine(Heading(1, content.metadata.header));
            body.AppendLine(buildMetadataTable());
            body.AppendLine(Heading(2, content.overview.header));
            body.AppendLine(Paragraph(content.overview.infoText));
            body.AppendLine(ParagraphRaw(Image("Flow Overview Diagram", content.overview.svgFile)));
            body.AppendLine(Heading(2, content.details.header));
            body.AppendLine(Paragraph(content.details.infoText));
            body.AppendLine(ParagraphRaw(Image(content.details.header, content.details.imageFileName + ".svg")));
            SaveHtmlFile(Path.Combine(content.folderPath, mainFileName),
                WrapInHtmlPage(content.metadata.header, body.ToString(), getNavigationHtml()));
        }

        private string getLinkFromAction(string name, bool fromSubfolder = false)
        {
            string prefix = fromSubfolder ? "" : "actions/";
            return (prefix + CharsetHelper.GetSafeName(name) + ".html").Replace(" ", "-");
        }
    }
}
