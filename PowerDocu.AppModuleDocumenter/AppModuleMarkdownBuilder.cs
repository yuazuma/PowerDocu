using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerDocu.Common;
using Grynwald.MarkdownGenerator;

namespace PowerDocu.AppModuleDocumenter
{
    class AppModuleMarkdownBuilder : MarkdownBuilder
    {
        private readonly AppModuleDocumentationContent content;
        private readonly string mainDocumentFileName;
        private readonly MdDocument mainDocument;
        private readonly DocumentSet<MdDocument> set;

        public AppModuleMarkdownBuilder(AppModuleDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            mainDocumentFileName = ("mda-" + content.filename + ".md").Replace(" ", "-");
            set = new DocumentSet<MdDocument>();
            mainDocument = set.CreateMdDocument(mainDocumentFileName);

            addOverview();
            addSecurityRoles();
            addNavigation();
            addTables();
            addViews();
            addCustomPages();
            addAppSettings();

            set.Save(content.folderPath);
            NotificationHelper.SendNotification("Created Markdown documentation for Model-Driven App: " + content.appModule.GetDisplayName());
        }

        private void addOverview()
        {
            mainDocument.Root.Add(new MdHeading(content.appModule.GetDisplayName(), 1));

            if (content.context?.Solution != null)
            {
                if (content.context?.Config?.documentSolution == true)
                    mainDocument.Root.Add(new MdParagraph(new MdCompositeSpan(new MdTextSpan("Solution: "), new MdLinkSpan(content.context.Solution.UniqueName, "../" + CrossDocLinkHelper.GetSolutionDocMdPath(content.context.Solution.UniqueName)))));
                else
                    mainDocument.Root.Add(new MdParagraph(new MdTextSpan("Solution: " + content.context.Solution.UniqueName)));
            }

            List<MdTableRow> tableRows = new List<MdTableRow>
            {
                new MdTableRow("Unique Name", content.appModule.UniqueName),
                new MdTableRow("Display Name", content.appModule.GetDisplayName()),
                new MdTableRow("Description", content.appModule.GetDescription()),
                new MdTableRow("Version", content.appModule.IntroducedVersion),
                new MdTableRow("Status", content.appModule.IsActive() ? "Active" : "Inactive"),
                new MdTableRow("Form Factor", content.appModule.GetFormFactorDisplayName()),
                new MdTableRow("Client Type", content.appModule.GetClientTypeDisplayName()),
                new MdTableRow(content.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion())
            };
            mainDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));

            // Descriptions in all languages
            if (content.appModule.Descriptions.Count > 1)
            {
                mainDocument.Root.Add(new MdHeading("Descriptions", 2));
                List<MdTableRow> descRows = new List<MdTableRow>();
                foreach (var desc in content.appModule.Descriptions)
                {
                    descRows.Add(new MdTableRow(desc.Key, desc.Value));
                }
                mainDocument.Root.Add(new MdTable(new MdTableRow("Language Code", "Description"), descRows));
            }
        }

        private void addSecurityRoles()
        {
            if (content.appModule.SecurityRoleIds.Count == 0) return;

            mainDocument.Root.Add(new MdHeading(content.headerSecurityRoles, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This app has {content.appModule.SecurityRoleIds.Count} security role(s) assigned.")));

            bool canLinkSolution = content.context?.Config?.documentSolution == true && content.context?.Solution != null;
            string solutionMdPath = canLinkSolution ? CrossDocLinkHelper.GetSolutionDocMdPath(content.context.Solution.UniqueName) : null;

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (string roleId in content.appModule.SecurityRoleIds)
            {
                string roleName = content.GetRoleNameById(roleId);
                MdSpan roleCell;
                if (canLinkSolution)
                {
                    string anchor = CrossDocLinkHelper.GetSolutionRoleMdAnchor(roleName, roleId);
                    roleCell = new MdLinkSpan(roleName, "../" + solutionMdPath + anchor);
                }
                else
                {
                    roleCell = new MdTextSpan(roleName);
                }
                tableRows.Add(new MdTableRow(roleCell, new MdTextSpan(roleId)));
            }
            mainDocument.Root.Add(new MdTable(new MdTableRow("Role Name", "Role ID"), tableRows));
        }

        private void addNavigation()
        {
            if (content.appModule.SiteMap == null) return;

            mainDocument.Root.Add(new MdHeading(content.headerNavigation, 2));

            var siteMap = content.appModule.SiteMap;

            // SiteMap properties
            List<MdTableRow> siteMapProps = new List<MdTableRow>
            {
                new MdTableRow("Show Home", siteMap.ShowHome.ToString()),
                new MdTableRow("Show Pinned", siteMap.ShowPinned.ToString()),
                new MdTableRow("Show Recents", siteMap.ShowRecents.ToString()),
                new MdTableRow("Collapsible Groups", siteMap.EnableCollapsibleGroups.ToString())
            };
            mainDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), siteMapProps));

            // SVG image link
            string svgFileName = "sitemap-" + content.filename.Replace(" ", "-") + ".svg";
            if (File.Exists(Path.Combine(content.folderPath, svgFileName)))
            {
                mainDocument.Root.Add(new MdParagraph(new MdImageSpan("Site Map", svgFileName)));
            }

            // Areas > Groups > SubAreas
            foreach (var area in siteMap.Areas)
            {
                mainDocument.Root.Add(new MdHeading($"Area: {area.Title}", 3));

                foreach (var group in area.Groups)
                {
                    mainDocument.Root.Add(new MdHeading($"Group: {group.Title}", 4));

                    if (group.SubAreas.Count > 0)
                    {
                        bool canLinkSolutionNav = content.context?.Config?.documentSolution == true && content.context?.Solution != null;
                        string solutionMdNav = canLinkSolutionNav ? CrossDocLinkHelper.GetSolutionDocMdPath(content.context.Solution.UniqueName) : null;

                        List<MdTableRow> subAreaRows = new List<MdTableRow>();
                        foreach (var subArea in group.SubAreas)
                        {
                            string title = !string.IsNullOrEmpty(subArea.Title) ? subArea.Title : subArea.Id;
                            string target = subArea.GetTargetDescription();
                            MdSpan targetCell;
                            if (canLinkSolutionNav && !string.IsNullOrEmpty(subArea.Entity))
                            {
                                string displayName = content.GetTableDisplayName(subArea.Entity);
                                string anchor = CrossDocLinkHelper.GetSolutionTableMdAnchor(displayName, subArea.Entity);
                                targetCell = new MdLinkSpan(target, "../" + solutionMdNav + anchor);
                            }
                            else
                            {
                                targetCell = new MdTextSpan(target);
                            }
                            subAreaRows.Add(new MdTableRow(new MdTextSpan(title), targetCell));
                        }
                        mainDocument.Root.Add(new MdTable(new MdTableRow("Navigation Item", "Target"), subAreaRows));
                    }
                }
            }
        }

        private void addTables()
        {
            var tables = content.appModule.GetTables();
            if (tables.Count == 0) return;

            mainDocument.Root.Add(new MdHeading(content.headerTables, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This app includes {tables.Count} table(s).")));

            bool canLinkSolution = content.context?.Config?.documentSolution == true && content.context?.Solution != null;
            string solutionMdPath = canLinkSolution ? CrossDocLinkHelper.GetSolutionDocMdPath(content.context.Solution.UniqueName) : null;

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (var comp in tables.OrderBy(c => c.SchemaName))
            {
                string displayName = content.GetTableDisplayName(comp.SchemaName);
                MdSpan nameCell;
                if (canLinkSolution)
                {
                    string anchor = CrossDocLinkHelper.GetSolutionTableMdAnchor(displayName, comp.SchemaName);
                    nameCell = new MdLinkSpan(displayName, "../" + solutionMdPath + anchor);
                }
                else
                {
                    nameCell = new MdTextSpan(displayName);
                }
                tableRows.Add(new MdTableRow(nameCell, new MdTextSpan(comp.SchemaName)));
            }
            mainDocument.Root.Add(new MdTable(new MdTableRow("Display Name", "Schema Name"), tableRows));
        }

        private void addViews()
        {
            var views = content.appModule.GetViews();
            if (views.Count == 0) return;

            mainDocument.Root.Add(new MdHeading(content.headerViews, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This app includes {views.Count} view(s).")));

            bool canLinkSolution = content.context?.Config?.documentSolution == true && content.context?.Solution != null;
            string solutionMdPath = canLinkSolution ? CrossDocLinkHelper.GetSolutionDocMdPath(content.context.Solution.UniqueName) : null;

            List<MdTableRow> tableRows = new List<MdTableRow>();
            var viewDetails = views.Select(comp => (comp, details: content.GetViewDetails(comp.ID)))
                .OrderBy(v => v.details.TableName).ThenBy(v => v.details.ViewName);
            foreach (var (comp, details) in viewDetails)
            {
                MdSpan tableCell;
                if (canLinkSolution && !string.IsNullOrEmpty(details.TableName))
                {
                    TableEntity tableEntity = content.allTables.FirstOrDefault(t =>
                        (t.getLocalizedName() ?? "").Equals(details.TableName, System.StringComparison.OrdinalIgnoreCase));
                    if (tableEntity != null)
                    {
                        string anchor = CrossDocLinkHelper.GetSolutionTableMdAnchor(tableEntity.getLocalizedName(), tableEntity.getName());
                        tableCell = new MdLinkSpan(details.TableName, "../" + solutionMdPath + anchor);
                    }
                    else
                    {
                        tableCell = new MdTextSpan(details.TableName);
                    }
                }
                else
                {
                    tableCell = new MdTextSpan(details.TableName);
                }
                tableRows.Add(new MdTableRow(tableCell, new MdTextSpan(details.ViewName), new MdTextSpan(details.QueryType), new MdTextSpan(comp.ID)));
            }
            mainDocument.Root.Add(new MdTable(new MdTableRow("Table", "View", "Query Type", "ID"), tableRows));
        }

        private void addCustomPages()
        {
            var customPages = content.appModule.GetCustomPages();
            if (customPages.Count == 0) return;

            mainDocument.Root.Add(new MdHeading(content.headerCustomPages, 2));
            mainDocument.Root.Add(new MdParagraph(new MdTextSpan($"This app includes {customPages.Count} custom page(s) (embedded canvas apps).")));

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (var page in customPages)
            {
                string displayName = content.GetCustomPageDisplayName(page);
                AppEntity app = content.GetCanvasAppForPage(page);
                string canvasAppCell;
                if (app != null && content.context?.Config?.documentApps == true)
                {
                    string safeFilename = CharsetHelper.GetSafeName(app.Name);
                    string indexFile = ("index-" + safeFilename + ".md").Replace(" ", "-");
                    string href = content.GetCanvasAppDocRelativePath(app, indexFile).Replace(" ", "%20");
                    canvasAppCell = $"[{app.Name}]({href})";
                }
                else
                {
                    canvasAppCell = page.CanvasAppName;
                }
                tableRows.Add(new MdTableRow(displayName, page.UniqueName, canvasAppCell));
            }
            mainDocument.Root.Add(new MdTable(new MdTableRow("Display Name", "Unique Name", "Canvas App"), tableRows));
        }

        private void addAppSettings()
        {
            if (content.appModule.AppSettings.Count == 0) return;

            mainDocument.Root.Add(new MdHeading(content.headerAppSettings, 2));

            List<MdTableRow> tableRows = new List<MdTableRow>();
            foreach (var setting in content.appModule.AppSettings)
            {
                tableRows.Add(new MdTableRow(setting.SettingName, setting.Value, setting.IsCustomizable ? "Yes" : "No"));
            }
            mainDocument.Root.Add(new MdTable(new MdTableRow("Setting", "Value", "Customizable"), tableRows));
        }
    }
}
