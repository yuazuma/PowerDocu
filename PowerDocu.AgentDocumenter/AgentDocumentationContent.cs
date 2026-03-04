
using System.IO;
using PowerDocu.Common;

namespace PowerDocu.AgentDocumenter
{

    class AgentDocumentationContent
    {
        public string folderPath,
           filename;
        public AgentEntity agent;

        public string headerDocumentationGenerated = "Documentation generated at";
        public string Details = "Details";
        public string Description = "Description";
        public string Orchestration = "Orchestration";
        public string OrchestrationText = "Use generative AI to determine how best to respond to users and events.";
        public string ResponseModel = "Response model";
        public string Instructions = "Instructions";
        public string Knowledge = "Knowledge";
        public string WebSearch = "Web Search";
        public string Tools = "Tools";
        public string Triggers = "Triggers";
        public string Agents = "Agents";
        public string Topics = "Topics";
        public string Entities = "Entities";
        public string Variables = "Variables";
        public string SuggestedPrompts = "Suggested prompts";
        public string SuggestedPromptsText = "Suggest ways of starting conversations for Teams and Microsoft 365 channels.";
        public AgentDocumentationContent(AgentEntity agent, string path)
        {
            NotificationHelper.SendNotification("Preparing documentation content for " + agent.Name);
            folderPath = path + CharsetHelper.GetSafeName(@"\AgentDoc " + agent.Name + @"\");
            Directory.CreateDirectory(folderPath);
            filename = CharsetHelper.GetSafeName(agent.Name);
            this.agent = agent;
        }

    }
}