using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PowerDocu.Common;

namespace PowerDocu.AIModelDocumenter
{
    class AIModelWordDocBuilder : WordDocBuilder
    {
        private readonly AIModelDocumentationContent content;

        public AIModelWordDocBuilder(AIModelDocumentationContent contentDocumentation, string template)
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
                addInstructions();
                addInputs();
                addModelResponse();
            }
            NotificationHelper.SendNotification("Created Word documentation for AI Model: " + content.aiModel.getName());
        }

        private void addOverview()
        {
            AddHeading(content.aiModel.getName(), "Heading1");
            body.AppendChild(new Paragraph(new Run()));

            Table table = CreateTable();
            table.Append(CreateRow(new Text("Name"), new Text(content.aiModel.getName())));
            if (content.context?.Solution != null)
            {
                table.Append(CreateRow(new Text("Solution"), new Text(content.context.Solution.UniqueName)));
            }
            table.Append(CreateRow(new Text("ID"), new Text(content.aiModel.getID())));
            table.Append(CreateRow(new Text("Template ID"), new Text(content.aiModel.getTemplateId())));
            table.Append(CreateRow(new Text("Documentation generated at"),
                new Text(PowerDocuReleaseHelper.GetTimestampWithVersion())));
            body.Append(table);
            body.AppendChild(new Paragraph(new Run()));
        }

        private void addInstructions()
        {
            AddHeading("Instructions", "Heading2");
            Table instructionsTable = CreateTable();
            Paragraph paraTable = new Paragraph();
            Run tableRun = paraTable.AppendChild(new Run());
            string prompt = content.aiModel.getPrompt();
            string[] promptLines = prompt.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            foreach (string promptLine in promptLines)
            {
                tableRun.AppendChild(new Text(promptLine));
                tableRun.AppendChild(new Break());
            }
            instructionsTable.Append(CreateRow(paraTable));
            body.Append(instructionsTable);
            body.AppendChild(new Paragraph(new Run()));
        }

        private void addInputs()
        {
            AddHeading("Inputs", "Heading2");
            List<AIModelInput> inputs = content.aiModel.getInputs().OrderBy(i => i.Text).ToList();
            if (inputs.Count > 0)
            {
                Table inputsTable = CreateTable();
                inputsTable.Append(CreateHeaderRow(
                    new Text("Name"),
                    new Text("ID"),
                    new Text("Type"),
                    new Text("Quick Text Value")
                ));

                foreach (AIModelInput input in inputs)
                {
                    inputsTable.Append(CreateRow(
                        new Text(input.Text ?? ""),
                        new Text(input.Id ?? ""),
                        new Text(input.Type ?? ""),
                        new Text(input.QuickTextValue ?? "")
                    ));
                }
                body.Append(inputsTable);
            }
            else
            {
                body.AppendChild(new Paragraph(new Run(new Text("No inputs defined."))));
            }
            body.AppendChild(new Paragraph(new Run()));
        }

        private void addModelResponse()
        {
            AddHeading("Model Response", "Heading2");
            AIModelOutput output = content.aiModel.getOutput();
            if (output != null)
            {
                Table outputTable = CreateTable();
                outputTable.Append(CreateRow(new Text("Output Formats"), new Text(string.Join(", ", output.Formats))));
                if (!String.IsNullOrEmpty(output.jsonSchema))
                {
                    outputTable.Append(CreateRow(new Text("Schema"), CreateRunWithLinebreaks(JsonUtil.JsonPrettify(output.jsonSchema))));
                    outputTable.Append(CreateRow(new Text("Examples"), CreateRunWithLinebreaks(JsonUtil.JsonPrettify(output.jsonExamples))));
                }
                body.Append(outputTable);
            }
            else
            {
                body.AppendChild(new Paragraph(new Run(new Text("No model response defined."))));
            }
            body.AppendChild(new Paragraph(new Run()));
        }
    }
}
