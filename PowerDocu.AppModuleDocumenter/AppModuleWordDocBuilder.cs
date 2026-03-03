using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerDocu.Common;

namespace PowerDocu.AppModuleDocumenter
{
    class AppModuleWordDocBuilder : WordDocBuilder
    {
        private readonly AppModuleDocumentationContent content;

        public AppModuleWordDocBuilder(AppModuleDocumentationContent contentDocumentation, string template)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            string filename = InitializeWordDocument(content.folderPath + content.filename, template);
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true))
            {
                mainPart = wordDocument.MainDocumentPart;
                body = mainPart.Document.Body;
                PrepareDocument(!String.IsNullOrEmpty(template));

                addOverview();
                addSecurityRoles();
                addNavigation();
                addTables();
                addViews();
                addCustomPages();
                addAppSettings();
            }
            NotificationHelper.SendNotification("Created Word documentation for Model-Driven App: " + content.appModule.GetDisplayName());
        }

        private void addOverview()
        {
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.appModule.GetDisplayName()));
            ApplyStyleToParagraph("Heading1", para);
            body.AppendChild(new Paragraph(new Run()));

            Table table = CreateTable();
            table.Append(CreateRow(new Text("Unique Name"), new Text(content.appModule.UniqueName)));
            table.Append(CreateRow(new Text("Display Name"), new Text(content.appModule.GetDisplayName())));
            table.Append(CreateRow(new Text("Description"), new Text(content.appModule.GetDescription())));
            table.Append(CreateRow(new Text("Version"), new Text(content.appModule.IntroducedVersion ?? "")));
            table.Append(CreateRow(new Text("Status"), new Text(content.appModule.IsActive() ? "Active" : "Inactive")));
            table.Append(CreateRow(new Text("Form Factor"), new Text(content.appModule.GetFormFactorDisplayName())));
            table.Append(CreateRow(new Text("Client Type"), new Text(content.appModule.GetClientTypeDisplayName())));
            table.Append(CreateRow(new Text(content.headerDocumentationGenerated),
                new Text(PowerDocuReleaseHelper.GetTimestampWithVersion())));
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // Descriptions in all languages
            if (content.appModule.Descriptions.Count > 1)
            {
                para = body.AppendChild(new Paragraph());
                run = para.AppendChild(new Run());
                run.AppendChild(new Text("Descriptions"));
                ApplyStyleToParagraph("Heading2", para);
                body.AppendChild(new Paragraph(new Run()));

                table = CreateTable();
                table.Append(CreateHeaderRow(new Text("Language Code"), new Text("Description")));
                foreach (var desc in content.appModule.Descriptions)
                {
                    table.Append(CreateRow(new Text(desc.Key), new Text(desc.Value)));
                }
                body.Append(table);
                body.AppendChild(new Paragraph(new Run(new Break())));
            }
        }

        private void addSecurityRoles()
        {
            if (content.appModule.SecurityRoleIds.Count == 0) return;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.headerSecurityRoles));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(
                new Text($"This app has {content.appModule.SecurityRoleIds.Count} security role(s) assigned."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Role Name"), new Text("Role ID")));
            foreach (string roleId in content.appModule.SecurityRoleIds)
            {
                string roleName = content.GetRoleNameById(roleId);
                table.Append(CreateRow(new Text(roleName), new Text(roleId)));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addNavigation()
        {
            if (content.appModule.SiteMap == null) return;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.headerNavigation));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run()));

            var siteMap = content.appModule.SiteMap;

            // SiteMap properties table
            Table propsTable = CreateTable();
            propsTable.Append(CreateRow(new Text("Show Home"), new Text(siteMap.ShowHome.ToString())));
            propsTable.Append(CreateRow(new Text("Show Pinned"), new Text(siteMap.ShowPinned.ToString())));
            propsTable.Append(CreateRow(new Text("Show Recents"), new Text(siteMap.ShowRecents.ToString())));
            propsTable.Append(CreateRow(new Text("Collapsible Groups"), new Text(siteMap.EnableCollapsibleGroups.ToString())));
            body.Append(propsTable);
            body.AppendChild(new Paragraph(new Run(new Break())));

            // SVG image
            string svgFileName = "sitemap-" + content.filename.Replace(" ", "-") + ".svg";
            string svgFilePath = Path.Combine(content.folderPath, svgFileName);
            if (File.Exists(svgFilePath) && SVGImages != null)
            {
                SVGImages[svgFileName] = svgFilePath;
            }

            // Areas > Groups > SubAreas
            foreach (var area in siteMap.Areas)
            {
                para = body.AppendChild(new Paragraph());
                run = para.AppendChild(new Run());
                run.AppendChild(new Text($"Area: {area.Title}"));
                ApplyStyleToParagraph("Heading3", para);

                foreach (var group in area.Groups)
                {
                    para = body.AppendChild(new Paragraph());
                    run = para.AppendChild(new Run());
                    run.AppendChild(new Text($"Group: {group.Title}"));
                    ApplyStyleToParagraph("Heading4", para);

                    if (group.SubAreas.Count > 0)
                    {
                        Table subAreaTable = CreateTable();
                        subAreaTable.Append(CreateHeaderRow(new Text("Navigation Item"), new Text("Target")));
                        foreach (var subArea in group.SubAreas)
                        {
                            string title = !string.IsNullOrEmpty(subArea.Title) ? subArea.Title : subArea.Id;
                            string target = subArea.GetTargetDescription();
                            subAreaTable.Append(CreateRow(new Text(title), new Text(target)));
                        }
                        body.Append(subAreaTable);
                        body.AppendChild(new Paragraph(new Run(new Break())));
                    }
                }
            }
        }

        private void addTables()
        {
            var tables = content.appModule.GetTables();
            if (tables.Count == 0) return;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.headerTables));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(
                new Text($"This app includes {tables.Count} table(s)."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Display Name"), new Text("Schema Name")));
            foreach (var comp in tables.OrderBy(c => c.SchemaName))
            {
                string displayName = content.GetTableDisplayName(comp.SchemaName);
                table.Append(CreateRow(new Text(displayName), new Text(comp.SchemaName)));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addViews()
        {
            var views = content.appModule.GetViews();
            if (views.Count == 0) return;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.headerViews));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(
                new Text($"This app includes {views.Count} view(s)."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("View"), new Text("ID")));
            foreach (var comp in views)
            {
                string name = !string.IsNullOrEmpty(comp.SchemaName) ? comp.SchemaName : comp.ID;
                table.Append(CreateRow(new Text(name), new Text(comp.ID)));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addCustomPages()
        {
            var customPages = content.appModule.GetCustomPages();
            if (customPages.Count == 0) return;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.headerCustomPages));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run(
                new Text($"This app includes {customPages.Count} custom page(s) (embedded canvas apps)."))));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Canvas App Page"), new Text("Customizable")));
            foreach (var page in customPages)
            {
                table.Append(CreateRow(new Text(page.CanvasAppName), new Text(page.IsCustomizable ? "Yes" : "No")));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }

        private void addAppSettings()
        {
            if (content.appModule.AppSettings.Count == 0) return;

            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(content.headerAppSettings));
            ApplyStyleToParagraph("Heading2", para);
            body.AppendChild(new Paragraph(new Run()));

            Table table = CreateTable();
            table.Append(CreateHeaderRow(new Text("Setting"), new Text("Value"), new Text("Customizable")));
            foreach (var setting in content.appModule.AppSettings)
            {
                table.Append(CreateRow(new Text(setting.SettingName), new Text(setting.Value), new Text(setting.IsCustomizable ? "Yes" : "No")));
            }
            body.Append(table);
            body.AppendChild(new Paragraph(new Run(new Break())));
        }
    }
}
