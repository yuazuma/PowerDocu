using System;
using System.Collections.Generic;
using System.IO;
using PowerDocu.Common;

namespace PowerDocu.AppModuleDocumenter
{
    public static class AppModuleDocumentationGenerator
    {
        /// <summary>
        /// Generates documentation for all Model-Driven Apps found in a solution zip file.
        /// Returns the list of parsed AppModuleEntity objects.
        /// </summary>
        public static List<AppModuleEntity> GenerateDocumentation(
            SolutionEntity solution,
            bool fullDocumentation,
            ConfigHelper config,
            string path
        )
        {
            if (solution?.Customizations == null)
            {
                NotificationHelper.SendNotification("No customizations found, skipping Model-Driven App documentation.");
                return new List<AppModuleEntity>();
            }

            DateTime startDocGeneration = DateTime.Now;
            List<AppModuleEntity> appModules = solution.Customizations.getAppModules();

            if (appModules == null || appModules.Count == 0)
            {
                NotificationHelper.SendNotification("No Model-Driven Apps found in the solution.");
                return new List<AppModuleEntity>();
            }

            NotificationHelper.SendNotification($"Found {appModules.Count} Model-Driven App(s) in the solution.");

            // Get cross-reference data from the solution
            List<RoleEntity> roles = solution.Customizations.getRoles();
            List<TableEntity> tables = solution.Customizations.getEntities();

            if (fullDocumentation)
            {
                foreach (AppModuleEntity appModule in appModules)
                {
                    AppModuleDocumentationContent content = new AppModuleDocumentationContent(appModule, path, roles, tables);

                    // Generate SiteMap SVG
                    if (appModule.SiteMap != null)
                    {
                        SiteMapSvgBuilder.GenerateSiteMapSvg(appModule.SiteMap, content.folderPath, content.filename);
                    }

                    if (config.outputFormat.Equals(OutputFormatHelper.Word) || config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Word documentation for Model-Driven App: " + appModule.GetDisplayName());
                        string wordTemplate = null;
                        if (!String.IsNullOrEmpty(config.wordTemplate) && File.Exists(config.wordTemplate))
                        {
                            wordTemplate = config.wordTemplate;
                        }
                        AppModuleWordDocBuilder wordDoc = new AppModuleWordDocBuilder(content, wordTemplate);
                    }
                    if (config.outputFormat.Equals(OutputFormatHelper.Markdown) || config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating Markdown documentation for Model-Driven App: " + appModule.GetDisplayName());
                        AppModuleMarkdownBuilder markdownDoc = new AppModuleMarkdownBuilder(content);
                    }
                    if (config.outputFormat.Equals(OutputFormatHelper.Html) || config.outputFormat.Equals(OutputFormatHelper.All))
                    {
                        NotificationHelper.SendNotification("Creating HTML documentation for Model-Driven App: " + appModule.GetDisplayName());
                        AppModuleHtmlBuilder htmlDoc = new AppModuleHtmlBuilder(content);
                    }
                }
            }

            DateTime endDocGeneration = DateTime.Now;
            NotificationHelper.SendNotification(
                $"AppModuleDocumenter: Processed {appModules.Count} Model-Driven App(s) in {(endDocGeneration - startDocGeneration).TotalSeconds} seconds."
            );
            return appModules;
        }
    }
}
