using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using PowerDocu.Common;

namespace PowerDocu.AppModuleDocumenter
{
    class AppModuleHtmlBuilder : HtmlBuilder
    {
        private readonly AppModuleDocumentationContent content;
        private readonly string mainFileName;

        public AppModuleHtmlBuilder(AppModuleDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            WriteDefaultStylesheet(content.folderPath);
            mainFileName = ("mda-" + content.filename + ".html").Replace(" ", "-");

            addOverviewPage();
            NotificationHelper.SendNotification("Created HTML documentation for Model-Driven App: " + content.appModule.GetDisplayName());
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
            navItemsList.AddRange(new (string label, string href)[]
            {
                ("Overview", "#overview"),
                ("Security Roles", "#security-roles"),
                ("Navigation", "#navigation"),
                ("Tables", "#tables"),
                ("Views", "#views"),
                ("Custom Pages", "#custom-pages"),
                ("App Settings", "#app-settings")
            });
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"<div class=\"nav-title\">{Encode(content.appModule.GetDisplayName())}</div>");
            sb.Append(NavigationList(navItemsList));
            return sb.ToString();
        }

        private void addOverviewPage()
        {
            StringBuilder body = new StringBuilder();

            // Overview
            body.AppendLine(HeadingWithId(1, content.appModule.GetDisplayName(), "overview"));

            body.Append(TableStart("Property", "Value"));
            body.Append(TableRow("Unique Name", content.appModule.UniqueName));
            body.Append(TableRow("Display Name", content.appModule.GetDisplayName()));
            body.Append(TableRow("Description", content.appModule.GetDescription()));
            body.Append(TableRow("Version", content.appModule.IntroducedVersion));
            body.Append(TableRow("Status", content.appModule.IsActive() ? "Active" : "Inactive"));
            body.Append(TableRow("Form Factor", content.appModule.GetFormFactorDisplayName()));
            body.Append(TableRow("Client Type", content.appModule.GetClientTypeDisplayName()));
            body.Append(TableRow(content.headerDocumentationGenerated, PowerDocuReleaseHelper.GetTimestampWithVersion()));
            body.AppendLine(TableEnd());

            // Descriptions in all languages
            if (content.appModule.Descriptions.Count > 1)
            {
                body.AppendLine(Heading(2, "Descriptions"));
                body.Append(TableStart("Language Code", "Description"));
                foreach (var desc in content.appModule.Descriptions)
                {
                    body.Append(TableRow(desc.Key, desc.Value));
                }
                body.AppendLine(TableEnd());
            }

            // Security Roles
            addSecurityRoles(body);

            // Navigation
            addNavigation(body);

            // Tables
            addTables(body);

            // Views
            addViews(body);

            // Custom Pages
            addCustomPages(body);

            // App Settings
            addAppSettings(body);

            SaveHtmlFile(Path.Combine(content.folderPath, mainFileName),
                WrapInHtmlPage($"Model-Driven App - {content.appModule.GetDisplayName()}", body.ToString(), getNavigationHtml()));
        }

        private void addSecurityRoles(StringBuilder body)
        {
            if (content.appModule.SecurityRoleIds.Count == 0) return;

            body.AppendLine(HeadingWithId(2, content.headerSecurityRoles, "security-roles"));
            body.AppendLine(Paragraph($"This app has {content.appModule.SecurityRoleIds.Count} security role(s) assigned."));

            body.Append(TableStart("Role Name", "Role ID"));
            foreach (string roleId in content.appModule.SecurityRoleIds)
            {
                string roleName = content.GetRoleNameById(roleId);
                if (content.context?.Config?.documentSolution == true && content.context?.Solution != null)
                {
                    string solutionHtml = CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName);
                    string anchor = CrossDocLinkHelper.GetSolutionRoleHtmlAnchor(roleName);
                    body.Append(TableRowRaw(Link(roleName, "../" + solutionHtml + anchor), Encode(roleId)));
                }
                else
                {
                    body.Append(TableRow(roleName, roleId));
                }
            }
            body.AppendLine(TableEnd());
        }

        private void addNavigation(StringBuilder body)
        {
            if (content.appModule.SiteMap == null) return;

            body.AppendLine(HeadingWithId(2, content.headerNavigation, "navigation"));

            var siteMap = content.appModule.SiteMap;

            body.Append(TableStart("Property", "Value"));
            body.Append(TableRow("Show Home", siteMap.ShowHome.ToString()));
            body.Append(TableRow("Show Pinned", siteMap.ShowPinned.ToString()));
            body.Append(TableRow("Show Recents", siteMap.ShowRecents.ToString()));
            body.Append(TableRow("Collapsible Groups", siteMap.EnableCollapsibleGroups.ToString()));
            body.AppendLine(TableEnd());

            // SVG image
            string svgFileName = "sitemap-" + content.filename.Replace(" ", "-") + ".svg";
            if (File.Exists(Path.Combine(content.folderPath, svgFileName)))
            {
                body.AppendLine(ImageWithClass("Site Map", svgFileName, "sitemap-diagram"));
            }

            // Areas > Groups > SubAreas
            foreach (var area in siteMap.Areas)
            {
                body.AppendLine(Heading(3, $"Area: {area.Title}"));

                foreach (var group in area.Groups)
                {
                    body.AppendLine(Heading(4, $"Group: {group.Title}"));

                    if (group.SubAreas.Count > 0)
                    {
                        body.Append(TableStart("Navigation Item", "Target"));
                        bool canLinkSolutionNav = content.context?.Config?.documentSolution == true && content.context?.Solution != null;
                        string solutionHtmlNav = canLinkSolutionNav ? CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName) : null;
                        foreach (var subArea in group.SubAreas)
                        {
                            string title = !string.IsNullOrEmpty(subArea.Title) ? subArea.Title : subArea.Id;
                            string target = subArea.GetTargetDescription();
                            // If the sub-area targets an entity/table, link to its solution documentation
                            if (canLinkSolutionNav && !string.IsNullOrEmpty(subArea.Entity))
                            {
                                string anchor = CrossDocLinkHelper.GetSolutionTableHtmlAnchor(subArea.Entity);
                                body.Append(TableRowRaw(Encode(title), Link(target, "../" + solutionHtmlNav + anchor)));
                            }
                            else
                            {
                                body.Append(TableRow(title, target));
                            }
                        }
                        body.AppendLine(TableEnd());
                    }
                }
            }
        }

        private void addTables(StringBuilder body)
        {
            var tables = content.appModule.GetTables();
            if (tables.Count == 0) return;

            body.AppendLine(HeadingWithId(2, content.headerTables, "tables"));
            body.AppendLine(Paragraph($"This app includes {tables.Count} table(s)."));

            body.Append(TableStart("Display Name", "Schema Name"));
            foreach (var comp in tables.OrderBy(c => c.SchemaName))
            {
                string displayName = content.GetTableDisplayName(comp.SchemaName);
                if (content.context?.Config?.documentSolution == true && content.context?.Solution != null)
                {
                    string solutionHtml = CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName);
                    string anchor = CrossDocLinkHelper.GetSolutionTableHtmlAnchor(comp.SchemaName);
                    body.Append(TableRowRaw(Link(displayName, "../" + solutionHtml + anchor), Encode(comp.SchemaName)));
                }
                else
                {
                    body.Append(TableRow(displayName, comp.SchemaName));
                }
            }
            body.AppendLine(TableEnd());
        }

        private void addViews(StringBuilder body)
        {
            var views = content.appModule.GetViews();
            if (views.Count == 0) return;

            body.AppendLine(HeadingWithId(2, content.headerViews, "views"));
            body.AppendLine(Paragraph($"This app includes {views.Count} view(s)."));

            bool canLinkSolution = content.context?.Config?.documentSolution == true && content.context?.Solution != null;
            string solutionHtml = canLinkSolution ? CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName) : null;

            body.Append(TableStart("Table", "View", "Query Type", "ID"));
            var viewDetails = views.Select(comp => (comp, details: content.GetViewDetails(comp.ID)))
                .OrderBy(v => v.details.TableName).ThenBy(v => v.details.ViewName);
            foreach (var (comp, details) in viewDetails)
            {
                if (canLinkSolution && !string.IsNullOrEmpty(details.TableName))
                {
                    // Find the table schema name to build the anchor
                    var tableEntity = content.allTables.FirstOrDefault(t =>
                        (t.getLocalizedName() ?? t.getName()).Equals(details.TableName, System.StringComparison.OrdinalIgnoreCase));
                    if (tableEntity != null)
                    {
                        string anchor = CrossDocLinkHelper.GetSolutionTableHtmlAnchor(tableEntity.getName());
                        body.Append(TableRowRaw(
                            Link(details.TableName, "../" + solutionHtml + anchor),
                            Encode(details.ViewName), Encode(details.QueryType), Encode(comp.ID)));
                        continue;
                    }
                }
                body.Append(TableRow(details.TableName, details.ViewName, details.QueryType, comp.ID));
            }
            body.AppendLine(TableEnd());
        }

        private void addCustomPages(StringBuilder body)
        {
            var customPages = content.appModule.GetCustomPages();
            if (customPages.Count == 0) return;

            body.AppendLine(HeadingWithId(2, content.headerCustomPages, "custom-pages"));
            body.AppendLine(Paragraph($"This app includes {customPages.Count} custom page(s) (embedded canvas apps)."));

            body.Append(TableStart("Display Name", "Unique Name", "Canvas App"));
            foreach (var page in customPages)
            {
                string displayName = content.GetCustomPageDisplayName(page);
                AppEntity app = content.GetCanvasAppForPage(page);
                string canvasAppCell;
                if (app != null && content.context?.Config?.documentApps == true)
                {
                    string href = "../" + CrossDocLinkHelper.GetAppDocHtmlPath(app.Name);
                    canvasAppCell = Link(app.Name, href);
                }
                else
                {
                    canvasAppCell = Encode(app?.Name ?? page.CanvasAppName);
                }
                body.Append(TableRowRaw(Encode(displayName), Encode(page.UniqueName), canvasAppCell));
            }
            body.AppendLine(TableEnd());
        }

        private void addAppSettings(StringBuilder body)
        {
            if (content.appModule.AppSettings.Count == 0) return;

            body.AppendLine(HeadingWithId(2, content.headerAppSettings, "app-settings"));

            body.Append(TableStart("Setting", "Value", "Customizable"));
            foreach (var setting in content.appModule.AppSettings)
            {
                body.Append(TableRow(setting.SettingName, setting.Value, setting.IsCustomizable ? "Yes" : "No"));
            }
            body.AppendLine(TableEnd());
        }
    }
}
