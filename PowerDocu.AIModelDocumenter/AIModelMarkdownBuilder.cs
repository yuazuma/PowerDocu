using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PowerDocu.Common;
using Grynwald.MarkdownGenerator;

namespace PowerDocu.AIModelDocumenter
{
    class AIModelMarkdownBuilder : MarkdownBuilder
    {
        private readonly AIModelDocumentationContent content;
        private readonly string mainDocumentFileName;
        private readonly MdDocument mainDocument;

        public AIModelMarkdownBuilder(AIModelDocumentationContent contentDocumentation)
        {
            content = contentDocumentation;
            Directory.CreateDirectory(content.folderPath);
            mainDocumentFileName = ("aimodel-" + content.filename + ".md").Replace(" ", "-");
            mainDocument = new MdDocument();

            addOverview();
            addInstructions();
            addInputs();
            addModelResponse();

            mainDocument.Save(content.folderPath + "/" + mainDocumentFileName);
            NotificationHelper.SendNotification("Created Markdown documentation for AI Model: " + content.aiModel.getName());
        }

        private void addOverview()
        {
            mainDocument.Root.Add(new MdHeading(content.aiModel.getName(), 1));

            if (content.context?.Solution != null)
            {
                if (content.context?.Config?.documentSolution == true)
                    mainDocument.Root.Add(new MdParagraph(new MdCompositeSpan(new MdTextSpan("Solution: "), new MdLinkSpan(content.context.Solution.UniqueName, "../" + CrossDocLinkHelper.GetSolutionDocMdPath(content.context.Solution.UniqueName)))));
                else
                    mainDocument.Root.Add(new MdParagraph(new MdTextSpan("Solution: " + content.context.Solution.UniqueName)));
            }

            List<MdTableRow> tableRows = new List<MdTableRow>
            {
                new MdTableRow("Name", content.aiModel.getName()),
                new MdTableRow("ID", content.aiModel.getID()),
                new MdTableRow("Template ID", content.aiModel.getTemplateId()),
                new MdTableRow("Documentation generated at", PowerDocuReleaseHelper.GetTimestampWithVersion())
            };
            mainDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), tableRows));
        }

        private void addInstructions()
        {
            mainDocument.Root.Add(new MdHeading("Instructions", 2));
            string prompt = content.aiModel.getPrompt();
            List<MdTableRow> instructionRows = new List<MdTableRow>
            {
                new MdTableRow(prompt)
            };
            mainDocument.Root.Add(new MdTable(new MdTableRow("Prompt"), instructionRows));
        }

        private void addInputs()
        {
            mainDocument.Root.Add(new MdHeading("Inputs", 2));
            List<AIModelInput> inputs = content.aiModel.getInputs().OrderBy(i => i.Text).ToList();
            if (inputs.Count > 0)
            {
                List<MdTableRow> inputRows = new List<MdTableRow>();
                foreach (AIModelInput input in inputs)
                {
                    inputRows.Add(new MdTableRow(input.Text ?? "", input.Id ?? "", input.Type ?? "", input.QuickTextValue ?? ""));
                }
                mainDocument.Root.Add(new MdTable(new MdTableRow("Name", "ID", "Type", "Quick Text Value"), inputRows));
            }
            else
            {
                mainDocument.Root.Add(new MdParagraph(new MdTextSpan("No inputs defined.")));
            }
        }

        private void addModelResponse()
        {
            mainDocument.Root.Add(new MdHeading("Model Response", 2));
            AIModelOutput output = content.aiModel.getOutput();
            if (output != null)
            {
                List<MdTableRow> outputRows = new List<MdTableRow>
                {
                    new MdTableRow("Output Formats", string.Join(", ", output.Formats))
                };
                if (!String.IsNullOrEmpty(output.jsonSchema))
                {
                    outputRows.Add(new MdTableRow("Schema", JsonUtil.JsonPrettify(output.jsonSchema)));
                    outputRows.Add(new MdTableRow("Examples", JsonUtil.JsonPrettify(output.jsonExamples)));
                }
                mainDocument.Root.Add(new MdTable(new MdTableRow("Property", "Value"), outputRows));
            }
            else
            {
                mainDocument.Root.Add(new MdParagraph(new MdTextSpan("No model response defined.")));
            }
        }
    }
}
