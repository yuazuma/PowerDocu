using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using PowerDocu.Common;

namespace PowerDocu.AIModelDocumenter
{
    class AIModelHtmlBuilder : HtmlBuilder
    {
        private readonly AIModelDocumentationContent content;
        private readonly string mainFileName;

        public AIModelHtmlBuilder(AIModelDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            WriteDefaultStylesheet(content.folderPath);
            mainFileName = ("aimodel-" + content.filename + ".html").Replace(" ", "-");

            addOverviewPage();
            NotificationHelper.SendNotification("Created HTML documentation for AI Model: " + content.aiModel.getName());
        }

        private string getNavigationHtml()
        {
            var navItems = new List<(string label, string href, int level)>();
            if (content.context?.Solution != null)
            {
                if (content.context?.Config?.documentSolution == true)
                    navItems.Add(("Solution", "../" + CrossDocLinkHelper.GetSolutionDocHtmlPath(content.context.Solution.UniqueName), 0));
                else
                    navItems.Add((content.context.Solution.UniqueName, "", 0));
            }
            navItems.AddRange(new (string label, string href, int level)[]
            {
                ("Overview", "#overview", 0),
                ("Instructions", "#instructions", 0),
                ("Inputs", "#inputs", 0),
                ("Model Response", "#model-response", 0)
            });
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"<div class=\"nav-title\">{Encode(content.aiModel.getName())}</div>");
            sb.Append(NavigationList(navItems));
            return sb.ToString();
        }

        private void addOverviewPage()
        {
            StringBuilder body = new StringBuilder();

            body.AppendLine(HeadingWithId(1, content.aiModel.getName(), "overview"));
            body.Append(TableStart("Property", "Value"));
            body.Append(TableRow("Name", content.aiModel.getName()));
            body.Append(TableRow("ID", content.aiModel.getID()));
            body.Append(TableRow("Template ID", content.aiModel.getTemplateId()));
            body.Append(TableRow("Documentation generated at", PowerDocuReleaseHelper.GetTimestampWithVersion()));
            body.AppendLine(TableEnd());

            addInstructions(body);
            addInputs(body);
            addModelResponse(body);

            SaveHtmlFile(Path.Combine(content.folderPath, mainFileName),
                WrapInHtmlPage("AI Model - " + content.aiModel.getName(), body.ToString(), getNavigationHtml()));
        }

        private void addInstructions(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, "Instructions", "instructions"));
            string prompt = content.aiModel.getPrompt();
            body.AppendLine(ParagraphWithLinebreaks(prompt));
        }

        private void addInputs(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, "Inputs", "inputs"));
            List<AIModelInput> inputs = content.aiModel.getInputs().OrderBy(i => i.Text).ToList();
            if (inputs.Count > 0)
            {
                body.Append(TableStart("Name", "ID", "Type", "Quick Text Value"));
                foreach (AIModelInput input in inputs)
                {
                    body.Append(TableRow(input.Text ?? "", input.Id ?? "", input.Type ?? "", input.QuickTextValue ?? ""));
                }
                body.AppendLine(TableEnd());
            }
            else
            {
                body.AppendLine(Paragraph("No inputs defined."));
            }
        }

        private void addModelResponse(StringBuilder body)
        {
            body.AppendLine(HeadingWithId(2, "Model Response", "model-response"));
            AIModelOutput output = content.aiModel.getOutput();
            if (output != null)
            {
                body.Append(TableStart("Property", "Value"));
                body.Append(TableRow("Output Formats", string.Join(", ", output.Formats)));
                body.AppendLine(TableEnd());
                if (!String.IsNullOrEmpty(output.jsonSchema))
                {
                    body.AppendLine(Paragraph("Schema:"));
                    body.AppendLine(CodeBlock(JsonUtil.JsonPrettify(output.jsonSchema)));
                    body.AppendLine(Paragraph("Examples:"));
                    body.AppendLine(CodeBlock(JsonUtil.JsonPrettify(output.jsonExamples)));
                }
            }
            else
            {
                body.AppendLine(Paragraph("No model response defined."));
            }
        }
    }
}
