using System.IO;
using PowerDocu.Common;

namespace PowerDocu.DesktopFlowDocumenter
{
    public class DesktopFlowDocumentationContent
    {
        public string folderPath, filename;
        public DesktopFlowEntity flow;
        public DocumentationContext context;

        public string headerOverview = "Overview";
        public string headerActionSteps = "Action Steps";
        public string headerVariables = "Variables";
        public string headerControlFlow = "Control Flow";
        public string headerModules = "Modules";
        public string headerConnectors = "Connectors";
        public string headerEnvironmentVariables = "Environment Variables";
        public string headerProperties = "Properties";
        public string headerSubflows = "Subflows";
        public string headerDocumentationGenerated = "Documentation generated at";

        public DesktopFlowDocumentationContent(DesktopFlowEntity flow, string path, DocumentationContext context)
        {
            NotificationHelper.SendNotification("Preparing documentation content for Desktop Flow: " + flow.GetDisplayName());
            this.flow = flow;
            this.context = context;
            folderPath = path + CharsetHelper.GetSafeName(@"\DesktopFlowDoc " + flow.GetDisplayName() + @"\");
            Directory.CreateDirectory(folderPath);
            filename = CharsetHelper.GetSafeName(flow.GetDisplayName());
        }
    }
}
