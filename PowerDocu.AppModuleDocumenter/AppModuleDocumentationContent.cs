using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerDocu.Common;

namespace PowerDocu.AppModuleDocumenter
{
    public class AppModuleDocumentationContent
    {
        public string folderPath, filename;
        public AppModuleEntity appModule;

        // Section headers
        public string headerOverview = "Overview";
        public string headerSecurityRoles = "Security Roles";
        public string headerNavigation = "Navigation (Site Map)";
        public string headerTables = "Tables";
        public string headerViews = "Views";
        public string headerCustomPages = "Custom Pages";
        public string headerAppSettings = "App Settings";
        public string headerDocumentationGenerated = "Documentation generated at";

        // Cross-reference data from the solution
        public List<RoleEntity> allRoles;
        public List<TableEntity> allTables;

        public AppModuleDocumentationContent(AppModuleEntity appModule, string path, List<RoleEntity> roles = null, List<TableEntity> tables = null)
        {
            NotificationHelper.SendNotification("Preparing documentation content for Model-Driven App: " + appModule.GetDisplayName());
            this.appModule = appModule;
            folderPath = path + CharsetHelper.GetSafeName(@"\AppModuleDoc " + appModule.GetDisplayName() + @"\");
            Directory.CreateDirectory(folderPath);
            filename = CharsetHelper.GetSafeName(appModule.GetDisplayName());
            allRoles = roles ?? new List<RoleEntity>();
            allTables = tables ?? new List<TableEntity>();
        }

        /// <summary>
        /// Resolves a security role GUID to its display name using the roles parsed from the solution.
        /// </summary>
        public string GetRoleNameById(string roleId)
        {
            if (string.IsNullOrEmpty(roleId)) return roleId;
            RoleEntity role = allRoles.FirstOrDefault(r => r.ID != null && r.ID.Equals(roleId, System.StringComparison.OrdinalIgnoreCase));
            return role?.Name ?? roleId;
        }

        /// <summary>
        /// Resolves a table schema name to its display name using the tables parsed from the solution.
        /// </summary>
        public string GetTableDisplayName(string schemaName)
        {
            if (string.IsNullOrEmpty(schemaName)) return schemaName;
            TableEntity table = allTables.FirstOrDefault(t => t.getName().Equals(schemaName, System.StringComparison.OrdinalIgnoreCase));
            return table?.getLocalizedName() ?? schemaName;
        }
    }
}
