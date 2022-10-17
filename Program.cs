using Azure;
using Azure.AI.TextAnalytics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

AzureKeyCredential credentials = new AzureKeyCredential("<Your API Key>");
Uri endpointUri = new Uri("<Your API endpoint>");
TextAnalyticsClient textAnalyticsClient = new TextAnalyticsClient(endpointUri, credentials);
string filePath = "<absolute path to .docx file with PII>";

using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
{

    WordprocessingDocument? newDoc = doc.Clone() as WordprocessingDocument;
    doc.Close();
    List<Paragraph>? paragraphs = newDoc?.MainDocumentPart?.Document?.Body?.ChildElements?.OfType<Paragraph>().ToList();

    if (paragraphs is not null)
    {
        foreach (Paragraph p in paragraphs)
        {
            List<Run>? runs = p.ChildElements?.OfType<Run>().ToList();

            if (runs is not null)
            {
                foreach (Run r in runs)
                {

                    if (!String.IsNullOrWhiteSpace(r.InnerText))
                    {

                        PiiEntityCollection piiEntities = (await textAnalyticsClient.RecognizePiiEntitiesAsync(r.InnerText)).Value;

                        Console.WriteLine($"Redacted Text: {piiEntities.RedactedText}");

                        if (piiEntities.Count > 0)
                        {
                            r.RemoveAllChildren<Text>();
                            r.AppendChild(new Text(piiEntities.RedactedText) { Space = SpaceProcessingModeValues.Preserve });

                            foreach (PiiEntity entity in piiEntities)
                            {
                                Console.WriteLine($"  Text: {entity.Text}");
                                Console.WriteLine($"  Category: {entity.Category}");

                                Console.WriteLine($"  SubCategory: {(!string.IsNullOrEmpty(entity.SubCategory) ? entity.SubCategory : String.Empty)}");

                                Console.WriteLine($"  Confidence score: {entity.ConfidenceScore}");
                                Console.WriteLine("");
                            }
                        }
                    }
                }
            }
        }

    }

    if (newDoc is not null)
    {
        string[] strings = filePath.Split(".");
        string ext = strings[strings.Length - 1];
        strings = strings.SkipLast(1).ToArray();
        string root = String.Join(String.Empty, strings);

        newDoc.SaveAs($@"{root}_redacted.{ext}");
        newDoc.Close();
    }
}