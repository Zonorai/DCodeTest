using Newtonsoft.Json;
using Novacode;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;

namespace DMobiAnalysis.Models
{
    public class DocumentProcessor
    {
        private List<WorkInstructionTextItem> Images { get; set; } = new List<WorkInstructionTextItem>();
        private int ImagesAddedCount { get; set; } = 0;
        private int Score { get; set; } = 100;

        public void ProcessDocuments(string path)
        {
            var docs = new List<ConversionResult>();
            var exceptions = new List<string>();

            var fileEntries = Directory.EnumerateFiles(path, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith(".doc") || s.EndsWith(".docx"));
            //var fileEntries = Directory.GetFiles(path, "ECG_M01_HP44_Illumination Test_Millw.pdf.*", SearchOption.TopDirectoryOnly);

            foreach (string filePath in fileEntries)
            {
                try
                {
                    var attachment = new Attachment { FileBytes = File.ReadAllBytes(filePath), FileName = Path.GetFileName(filePath) };

                    var docResult = ConvertWorkInstructionsFromWordDoc(attachment);
                    docs.Add(docResult);

                    var resultTypeName = "Success";

                    if (docResult.Aborted)
                    {
                        resultTypeName = "Aborted";
                    }
                    else if (docResult.HasRuleViolations)
                    {
                        resultTypeName = "SuccessWithWarnings";
                    }

                    var outputFilePath = Path.Combine(path, resultTypeName, attachment.FileName);

                    System.IO.Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath));

                    File.WriteAllBytes(outputFilePath, attachment.FileBytes);

                    File.WriteAllText(outputFilePath + ".result.json", JsonConvert.SerializeObject(docResult, Newtonsoft.Json.Formatting.Indented));
                }

                catch (Exception ex)
                {
                    var msg = $"Error processing filename '{filePath}': {ex.Message}";
                    exceptions.Add(msg);
                }
            }

            Console.WriteLine($"Processed {docs.Count} docs. {docs.Count(d => !d.Aborted)} successful, {docs.Count(d => d.Aborted)} failed. {docs.Count(d => d.HasRuleViolations)} had warnings");
        }
        public ConversionResult ConvertWorkInstructionsFromWordDoc(Attachment attachment)
        {
            Score = 100;
            ImagesAddedCount = 0;
            Images = new List<WorkInstructionTextItem>();

            ConversionResult result = new ConversionResult();
            WorkInstruction workInstruction = new WorkInstruction();
            List<WorkInstructionTextItem> instructions = new List<WorkInstructionTextItem>();

            using (MemoryStream stream = new MemoryStream(attachment.FileBytes))
            {
                using (DocX document = DocX.Load(stream))
                {
                    Images = GetImagesFromDocument(document);
                    if(Images.Count > 0)
                    {
                        Score -= 20;
                    }
                    var tablesWithWorkInstructions = document.Tables.Where(table => DoesTableContainWork(table));
                    if (tablesWithWorkInstructions.Count() < 1)
                    {
                        result.AddRuleViolation("Tables Required", true, "This document contains no tables to process");
                        return result;
                    }
                    result.Filename = attachment.FileName;
                    //Instructions can span over multiple Tables*
                    //When they span they repeat headers
                    if (tablesWithWorkInstructions.Count() == 1)
                    {
                        if (Images.Count > 0)
                        {
                            try
                            {
                                instructions = ExtractInstructionsFromTable(tablesWithWorkInstructions.First(), document);
                            }
                            catch (WhiteSpaceInstructions e)
                            {
                                result.AddRuleViolation("HasInstuctions", true, "Instructions found but they had no value");
                                return result;
                            }
                            catch (NoInstructions e)
                            {
                                result.AddRuleViolation("HasInstuctions", true, "Must contain instructions");
                                return result;
                            }
                        }
                        else
                        {
                            try
                            {
                                instructions = ExtractInstructionsFromTable(tablesWithWorkInstructions.First());
                            }
                            catch (WhiteSpaceInstructions e)
                            {
                                result.AddRuleViolation("HasInstuctions", true, "Instructions found but they had no value");
                                return result;
                            }
                            catch (NoInstructions e)
                            {
                                result.AddRuleViolation("HasInstuctions", true, "Must contain instructions");
                                return result;
                            }

                        }
                    }
                    else
                    {

                        foreach (var table in tablesWithWorkInstructions)
                        {
                            if (Images.Count > 0)
                            {
                                var instructionsExtracted = ExtractInstructionsFromTable(table, document);
                                if(instructionsExtracted.Count > 0)
                                {
                                    instructions.AddRange(instructionsExtracted);
                                }
                                
                            }
                            else
                            {
                                List<WorkInstructionTextItem> textItems = ExtractInstructionsFromTable(table);
                                if (textItems.Count > 0)
                                {
                                    instructions.AddRange(textItems);
                                }
                                
                            }
                        }

                    }
                    if (!(instructions.Count > 0))
                    {
                        result.AddRuleViolation("HasInstuctions", true, "Document must contain instructions");
                        return result;
                    }
                    else
                    {
                        workInstruction.InstructionsAsText = instructions;
                        workInstruction.SourceFilename = attachment.FileName;
                        result.WorkInstructions = workInstruction;
                        result.ConversionScore = Score;
                        return result;
                    }
                }
            }
        }
        private string GetGroupNameForInstructions(Table table)
        {
            var containsMagic = table.Paragraphs.Where(paragraph => paragraph.MagicText.Count > 0);
            foreach(var paragraph in containsMagic)
            {
                if(paragraph.MagicText.Count < 1)
                {
                    return string.Empty;
                }
                try
                {
                    if (paragraph.MagicText.Any(magic => magic.formatting.UnderlineStyle == UnderlineStyle.singleLine))
                    {
                        return paragraph.Text;
                    }
                }
                catch
                {

                }
            }
            return string.Empty;
        }
        private List<WorkInstructionTextItem> GetImagesFromDocument(DocX doc)
        {
            var images = new List<WorkInstructionTextItem>();
            if (doc.Images.Count < 1)
            {
                return images;
            }
            foreach (var image in doc.Images)
            {
                using (Stream m = image.GetStream(FileMode.Open, FileAccess.Read))
                {
                    using (MemoryStream memStream = new MemoryStream())
                    {
                        m.CopyTo(memStream);
                        byte[] imageBytes = memStream.ToArray();
                        string base64String = Convert.ToBase64String(imageBytes);

                        WorkInstructionTextItem textItem = new WorkInstructionTextItem { Text = base64String, WorkInstructionItemId = Guid.NewGuid() };
                        images.Add(textItem);
                    }

                }

            }
            return images;

        }
        private List<WorkInstructionTextItem> ExtractInstructionsFromTable(Table table, DocX doc)
        {
            if (table.Paragraphs.Count() < 1)
            {
                throw new Exception("Table has no instructions");
            }

            //
            List<Paragraph> onlyLists = new List<Paragraph>();

            onlyLists = table.Paragraphs.Where(paragraph => paragraph.IsListItem == true && (paragraph.ListItemType == ListItemType.Bulleted  || paragraph.ListItemType == ListItemType.Numbered)).ToList();

            if (onlyLists.Count < 1)
            {
                Score -= 50;
                //Maybe we find a manual numbered list
                if(table.Paragraphs.Any(text => text.Text.Contains("1.")))
                {
                    System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"\d{1,2}\.");
                    int indexOfFirst = table.Paragraphs.FindIndex(text => text.Text.Contains("1."));
                    var noWhiteSpace = table.Paragraphs.Where(parag => !string.IsNullOrEmpty(parag.Text));
                    var manualNumbered = noWhiteSpace.Where(parag => regex.IsMatch(parag.Text.TrimStart().Substring(0, 2)));
                    onlyLists = manualNumbered.ToList();
                }
            }
            //Remove and WhiteSpace Paragraphs
            var whiteSpaceRemoved = onlyLists.Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text)).ToList();

            var instructionsList = GetInstructionsTextFromParagraphs(whiteSpaceRemoved);

            int count = ImagesAddedCount > 0 ? ImagesAddedCount - 1 : 0;

            if (table.Xml.NextNode.ToString().Contains("graphicData"))
            {
                if (Images.Count - ImagesAddedCount > 0)
                {
                    instructionsList.Add(Images[count]);
                    ImagesAddedCount += 1;
                }

            }
            if (table.Xml.PreviousNode.ToString().Contains("graphicData"))
            {

                if (Images.Count - ImagesAddedCount > 0)
                {
                    instructionsList.Insert(0, Images[count]);
                    ImagesAddedCount += 1;
                }
            }

            instructionsList.ForEach(textItem => textItem.GroupName = GetGroupNameForInstructions(table));

            return instructionsList;

        }
        private List<WorkInstructionTextItem> ExtractInstructionsFromTable(Table table)
        {
            if (table.Paragraphs.Count() < 1)
            {
                throw new NoInstructions();
            }

            //
            var onlyBulletLists = table.Paragraphs.Where(paragraph => paragraph.IsListItem && paragraph.ListItemType == ListItemType.Bulleted).ToList();
            //Remove and WhiteSpace Paragraphs
            var whiteSpaceRemoved = onlyBulletLists.Where(paragraph => !string.IsNullOrEmpty(paragraph.Text)).ToList();
            if (whiteSpaceRemoved.Count < 1)
            {
                throw new WhiteSpaceInstructions();
            }
            var instructionsList = GetInstructionsTextFromParagraphs(whiteSpaceRemoved);

            instructionsList.ForEach(textItem => textItem.GroupName = GetGroupNameForInstructions(table));

            return instructionsList;

        }
        private List<WorkInstructionTextItem> GetInstructionsTextFromParagraphs(List<Paragraph> paragraphs)
        {

            List<WorkInstructionTextItem> textItems = new List<WorkInstructionTextItem>();
            foreach (var paragraph in paragraphs)
            {
                WorkInstructionTextItem textItem = new WorkInstructionTextItem();

                textItem.WorkInstructionItemId = Guid.NewGuid();
                textItem.Text = paragraph.Text;
                textItems.Add(textItem);
            }

            return textItems;
        }
        private bool DoesTableContainWork(Table table)
        {
            if (table.Paragraphs.Any(paragraph => paragraph.Text.ToLower() == "work instruction" || paragraph.Text.ToLower() == "work instructions"))
            {
                return true;
            }
            if (table.Paragraphs.Any(paragraph => paragraph.Text.ToLower() == "tasks executed" || paragraph.Text.ToLower() == "additional work required"))
            {
                return true;
            }
            return false;
        }
        #region Models

        public class ConversionResult
        {
            public string Filename;

            public int ConversionScore;

            private List<RuleViolation> _ruleViolations = new List<RuleViolation>();

            public bool Aborted => _ruleViolations.Any((RuleViolation t) => t.IsCritical);

            public bool HasRuleViolations => _ruleViolations.Count > 0;

            public ReadOnlyCollection<RuleViolation> RuleViolations => new ReadOnlyCollection<RuleViolation>(_ruleViolations);

            public WorkInstruction WorkInstructions { get; set; } = new WorkInstruction();

            public void AddRuleViolation(string ruleName, bool abort, string warningText)
            {
                _ruleViolations.Add(new RuleViolation
                {
                    Message = warningText,
                    IsCritical = abort,
                    Rule = ruleName
                });
            }
        }
        public class WorkInstruction
        {
            public List<WorkInstructionTextItem> InstructionsAsText { get; set; }
            public string SourceFilename { get; set; }
        }

        public class RuleViolation
        {
            public string Rule { get; set; }
            public string Message { get; set; }
            public bool IsCritical { get; set; }
        }

        public class WorkInstructionTextItem
        {
            public Guid WorkInstructionItemId { get; set; }

            /// <summary>
            /// Items are grouped by <see cref="P:dCode.Models.PlantMaintenance.WorkInstructionTextItem.GroupName" /> when displayed as a Survey
            /// </summary>
            public string GroupName { get; set; }

            public string Text { get; set; }

            public string SubText { get; set; }
        }

        public class Attachment
        {
            public string Title { get; set; }

            public byte[] FileBytes { get; set; }

            public string FileName { get; set; }

            public string MimeType { get; set; }

            public string Path { get; set; }

            public string ExternalId { get; set; }

            public long Size { get; set; }

            public DateTime CreatedDate { get; set; }
        }
        #endregion Models
    }
}

