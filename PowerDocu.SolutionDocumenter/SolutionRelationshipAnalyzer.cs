using System;
using System.Collections.Generic;
using System.Linq;
using PowerDocu.Common;

namespace PowerDocu.SolutionDocumenter
{
    public static class SolutionRelationshipAnalyzer
    {
        public static (List<ComponentRelationship> Relationships, HashSet<SolutionComponentNode> AllComponents) Analyze(SolutionDocumentationContent content)
        {
            var relationships = new List<ComponentRelationship>();
            var allComponents = new HashSet<SolutionComponentNode>();

            CollectAllComponents(content, allComponents);
            CollectAgentRelationships(content, relationships);
            CollectFlowRelationships(content, relationships);
            CollectAppRelationships(content, relationships);
            CollectTableRelationships(content, relationships);
            CollectModelDrivenAppRelationships(content, relationships);

            return (relationships, allComponents);
        }

        private static void CollectAllComponents(SolutionDocumentationContent content, HashSet<SolutionComponentNode> allComponents)
        {
            foreach (var agent in content.agents)
            {
                allComponents.Add(new SolutionComponentNode("Agent", agent.Name));
            }
            foreach (var flow in content.flows)
            {
                allComponents.Add(new SolutionComponentNode("Flow", flow.Name));
            }
            foreach (var app in content.apps)
            {
                allComponents.Add(new SolutionComponentNode("Canvas App", app.Name));
            }
            foreach (var appModule in content.appModules)
            {
                allComponents.Add(new SolutionComponentNode("Model-Driven App", appModule.GetDisplayName()));
            }
            if (content.solution?.Customizations != null)
            {
                foreach (var table in content.solution.Customizations.getEntities())
                {
                    string name = table.getLocalizedName();
                    if (string.IsNullOrEmpty(name)) name = table.getName();
                    allComponents.Add(new SolutionComponentNode("Table", name));
                }
                foreach (var aiModel in content.solution.Customizations.getAIModels())
                {
                    allComponents.Add(new SolutionComponentNode("AI Model", aiModel.getName()));
                }
            }
            if (content.solution != null)
            {
                foreach (var envVar in content.solution.EnvironmentVariables)
                {
                    allComponents.Add(new SolutionComponentNode("Environment Variable", envVar.DisplayName ?? envVar.Name));
                }
                // Add remaining component types that aren't explicitly parsed
                foreach (var comp in content.solution.Components)
                {
                    string displayName = content.GetDisplayNameForComponent(comp);
                    switch (comp.Type)
                    {
                        case "Workflow":
                        case "Entity":
                        case "Canvas App":
                        case "AI Project":
                            // Already added above via parsed entities
                            break;
                        case "Connector":
                        case "Connection Reference":
                            allComponents.Add(new SolutionComponentNode("Connector", ConnectorHelper.ResolveConnectorDisplayName(displayName)));
                            break;
                        case "Role":
                            allComponents.Add(new SolutionComponentNode("Security Role", displayName));
                            break;
                        case "Option Set":
                            allComponents.Add(new SolutionComponentNode("Option Set", displayName));
                            break;
                        default:
                            allComponents.Add(new SolutionComponentNode(comp.Type, displayName));
                            break;
                    }
                }
            }
        }

        private static void CollectAgentRelationships(SolutionDocumentationContent content, List<ComponentRelationship> relationships)
        {
            foreach (var agent in content.agents)
            {
                // Agent → Flow (tools)
                foreach (var toolInfo in agent.GetAllToolInfos())
                {
                    if (!string.IsNullOrEmpty(toolInfo.FlowId))
                    {
                        string flowName = content.context?.GetFlowNameById(toolInfo.FlowId);
                        if (!string.IsNullOrEmpty(flowName))
                        {
                            relationships.Add(new ComponentRelationship("Agent", agent.Name, "Flow", flowName, "uses as tool"));
                        }
                    }
                    if (!string.IsNullOrEmpty(toolInfo.ConnectionReference))
                    {
                        string connectorName = ConnectorHelper.ResolveConnectorDisplayName(toolInfo.ConnectionReference);
                        relationships.Add(new ComponentRelationship("Agent", agent.Name, "Connector", connectorName, "uses connector"));
                    }
                }

                // Agent → Agent (connected agents)
                foreach (var connectedAgentInfo in agent.GetAllConnectedAgentInfos())
                {
                    string targetName = connectedAgentInfo.Name;
                    if (!string.IsNullOrEmpty(connectedAgentInfo.BotSchemaName) && content.context != null)
                    {
                        string resolved = content.context.GetAgentNameBySchemaName(connectedAgentInfo.BotSchemaName);
                        if (!string.IsNullOrEmpty(resolved)) targetName = resolved;
                    }
                    relationships.Add(new ComponentRelationship("Agent", agent.Name, "Agent", targetName, "invokes agent"));
                }

                // Agent → Table (knowledge sources)
                foreach (var dvTable in agent.DvTableSearchEntities)
                {
                    if (!string.IsNullOrEmpty(dvTable.EntityLogicalName))
                    {
                        string tableName = content.context?.GetTableDisplayName(dvTable.EntityLogicalName);
                        if (string.IsNullOrEmpty(tableName)) tableName = dvTable.EntityLogicalName;
                        relationships.Add(new ComponentRelationship("Agent", agent.Name, "Table", tableName, "knowledge source"));
                    }
                }

                // Agent → AI Model (via AI Plugin Operations)
                var addedAIModels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var operation in agent.AIPluginOperations)
                {
                    if (string.IsNullOrEmpty(operation.AIModelId))
                        continue;

                    string trimmedId = operation.AIModelId.Trim('{', '}');
                    var model = agent.AIModels.FirstOrDefault(m =>
                        m.getID().Trim('{', '}').Equals(trimmedId, StringComparison.OrdinalIgnoreCase));
                    if (model == null)
                        continue;

                    string aiModelName = model.getName();
                    if (!string.IsNullOrEmpty(aiModelName) && addedAIModels.Add(aiModelName))
                    {
                        relationships.Add(new ComponentRelationship("Agent", agent.Name, "AI Model", aiModelName, "uses AI model"));
                    }
                }
            }
        }

        private static void CollectFlowRelationships(SolutionDocumentationContent content, List<ComponentRelationship> relationships)
        {
            foreach (var flow in content.flows)
            {
                foreach (var connRef in flow.connectionReferences)
                {
                    if (!string.IsNullOrEmpty(connRef.Name))
                    {
                        string connectorName = ConnectorHelper.ResolveConnectorDisplayName(connRef.Name);
                        relationships.Add(new ComponentRelationship("Flow", flow.Name, "Connector", connectorName, "uses connector"));
                    }
                }
            }
        }

        private static void CollectAppRelationships(SolutionDocumentationContent content, List<ComponentRelationship> relationships)
        {
            // Build lookup of known table names for matching data sources to tables
            var tablesBySchemaName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var tablesByDisplayName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (content.solution?.Customizations != null)
            {
                foreach (var table in content.solution.Customizations.getEntities())
                {
                    string displayName = table.getLocalizedName();
                    if (string.IsNullOrEmpty(displayName)) displayName = table.getName();
                    tablesBySchemaName[table.getName()] = displayName;
                    if (!string.IsNullOrEmpty(table.getLocalizedName()))
                        tablesByDisplayName[table.getLocalizedName()] = displayName;
                }
            }

            foreach (var app in content.apps)
            {
                foreach (var ds in app.DataSources)
                {
                    //we skip sample data sources, auxiliary data sources (such as views), and collections
                    if (ds.isSampleDataSource()) continue;
                    if (ds.isAuxiliaryDataSource()) continue;
                    if (string.IsNullOrEmpty(ds.Name)) continue;
                    if (ds.Type == "CollectionDataSourceInfo") continue;
                    // Try to match this data source to a known Dataverse table
                    string matchedTableName = TryMatchDataSourceToTable(ds, tablesBySchemaName, tablesByDisplayName);
                    if (matchedTableName != null)
                    {
                        relationships.Add(new ComponentRelationship("Canvas App", app.Name, "Table", matchedTableName, "data source"));
                    }
                    else
                    {
                        relationships.Add(new ComponentRelationship("Canvas App", app.Name, "Data Source", ds.Name, "data source"));
                    }
                }
            }
        }

        private static string TryMatchDataSourceToTable(DataSource ds, Dictionary<string, string> tablesBySchemaName, Dictionary<string, string> tablesByDisplayName)
        {
            // 1) Check TableDefinition property for a Dataverse logical name (most reliable)
            var tableDefProp = ds.Properties.FirstOrDefault(p => p.expressionOperator == "TableDefinition");
            if (tableDefProp != null)
            {
                var info = TableDefinitionHelper.Parse(tableDefProp);
                if (info != null && !string.IsNullOrEmpty(info.LogicalName))
                {
                    if (tablesBySchemaName.TryGetValue(info.LogicalName, out string displayName))
                        return displayName;
                    // Table exists in the app but not in this solution — still a known Dataverse table
                    return info.DisplayName ?? info.LogicalName;
                }
            }

            // 2) Match data source name directly against table schema or display names
            if (tablesBySchemaName.TryGetValue(ds.Name, out string matched))
                return matched;
            if (tablesByDisplayName.TryGetValue(ds.Name, out matched))
                return matched;

            return null;
        }

        private static void CollectTableRelationships(SolutionDocumentationContent content, List<ComponentRelationship> relationships)
        {
            if (content.solution?.Customizations == null) return;

            var tables = content.solution.Customizations.getEntities();
            var entityRelationships = content.solution.Customizations.getEntityRelationships();

            foreach (var er in entityRelationships.Where(o => o.getRelationshipType().Equals("ManyToMany")))
            {
                string firstName = ResolveTableDisplayName(tables, er.getFirstEntityName());
                string secondName = ResolveTableDisplayName(tables, er.getSecondEntityName());
                relationships.Add(new ComponentRelationship("Table", firstName, "Table", secondName, "many-to-many"));
            }

            // One-to-many via lookup columns
            foreach (var table in tables)
            {
                string tableName = table.getLocalizedName();
                if (string.IsNullOrEmpty(tableName)) tableName = table.getName();

                foreach (var lookupCol in table.GetColumns().Where(c => c.isNonDefaultLookUpColumn()))
                {
                    var er = entityRelationships
                        .Where(o => o.getReferencingAttributeName().Equals(lookupCol.getLogicalName(), StringComparison.OrdinalIgnoreCase))
                        .FirstOrDefault();
                    if (er != null)
                    {
                        string referencedName = ResolveTableDisplayName(tables, er.getReferencedEntityName());
                        relationships.Add(new ComponentRelationship("Table", tableName, "Table", referencedName, "lookup"));
                    }
                }
            }
        }

        private static void CollectModelDrivenAppRelationships(SolutionDocumentationContent content, List<ComponentRelationship> relationships)
        {
            foreach (var appModule in content.appModules)
            {
                string appName = appModule.GetDisplayName();

                // Tables referenced via components
                foreach (var tableComp in appModule.GetTables())
                {
                    string tableName = content.context?.GetTableDisplayName(tableComp.SchemaName);
                    if (string.IsNullOrEmpty(tableName)) tableName = tableComp.SchemaName;
                    relationships.Add(new ComponentRelationship("Model-Driven App", appName, "Table", tableName, "uses table"));
                }

                // Tables referenced via sitemap sub-areas
                if (appModule.SiteMap != null)
                {
                    foreach (var area in appModule.SiteMap.Areas)
                    {
                        foreach (var group in area.Groups)
                        {
                            foreach (var subArea in group.SubAreas)
                            {
                                if (!string.IsNullOrEmpty(subArea.Entity))
                                {
                                    string tableName = content.context?.GetTableDisplayName(subArea.Entity);
                                    if (string.IsNullOrEmpty(tableName)) tableName = subArea.Entity;
                                    relationships.Add(new ComponentRelationship("Model-Driven App", appName, "Table", tableName, "navigates to"));
                                }
                            }
                        }
                    }
                }
            }
        }

        private static string ResolveTableDisplayName(List<TableEntity> tables, string schemaName)
        {
            var table = tables.Find(t => t.getName().Equals(schemaName, StringComparison.OrdinalIgnoreCase));
            if (table != null)
            {
                string localized = table.getLocalizedName();
                if (!string.IsNullOrEmpty(localized)) return localized;
            }
            return schemaName;
        }
    }
}
