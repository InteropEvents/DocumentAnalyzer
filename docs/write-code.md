# The code

- [Add the usings](#1-add-the-usings)
- [Create an Azure TextAnalyticsClient](#2-create-an-azure-textanalyticsclient)
- [Open the document and clone it](#3-open-the-document-and-clone-it)
- [Loop over the paragraphs](#4-loop-over-the-paragraphs)
- [Get a list of all the `Run` elements](#5-get-a-list-of-all-the-run-elements)
- [Check each run for PII and add the redacted text to the cloned document](#6-check-each-run-for-pii-and-add-the-redacted-text-to-the-cloned-document)
- [Save the new document with a new name](#7-save-the-new-document-with-a-new-name)

## 1. Add the usings

1. Open Program.cs and delete the contents.

2. Add the following usings to the top of the file:

```csharp
using Azure;
using Azure.AI.TextAnalytics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
```

## 2. Create an Azure TextAnalyticsClient

Below the usings, create an `AzureKeyCredential` and `Uri` with your API key and endpoint [from Azure](./setup.md) and use them to create the TextAnalyticsClient and store the file path in a variable.

```csharp
AzureKeyCredential credentials = new AzureKeyCredential("<Your API Key>");
Uri endpointUri = new Uri("<Your API endpoint>");
TextAnalyticsClient textAnalyticsClient = new TextAnalyticsClient(endpointUri, credentials);
string filePath = "<absolute path to .docx file with PII>";
```

## 3. Open the document and clone it

- First open the document using the [WordprocessingDocument.Open](https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.wordprocessingdocument.open) method and indicate that the document should open for read access only by passing false as the second parameter.

- Next use the [OpenXmlElement.Clone](https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.openxmlelement.clone?view=openxml-2.8.1) method to clone a copy of the document.

- Then close the original document.

- Finally create a list of the document body's paragraph elements.

```csharp
using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
{
    WordprocessingDocument? newDoc = doc.Clone() as WordprocessingDocument;
    doc.Close();
    List<Paragraph>? paragraphs = newDoc?.MainDocumentPart?.Document?.Body?.ChildElements?.OfType<Paragraph>().ToList();

    // Code for step 4 goes here
}
```

## 4. Loop over the paragraphs

- `paragraphs` could be null, so use the `is` operator to check for null

- Then loop over the `paragraphs` list.

```csharp
if (paragraphs is not null)
{
    foreach (Paragraph p in paragraphs)
    {
        // Code for step 5 goes here
    }
}
```

## 5. Get a list of all the `Run` elements

- Create a list of runs for each paragraph and store it in a variable called `runs`.

- The list could be null so use the `in` operator again to check for null before looping over `runs`

```csharp
List<Run>? runs = p.ChildElements?.OfType<Run>().ToList();

if (runs is not null)
{
    foreach (Run r in runs)
    {
        // Code for the step 6 goes here
    }
}
```

## 6. Check each run for PII and add the redacted text to the cloned document

- Check that `r.InnerText` is not null or whitespace, because those values cause the `TextAnalyticsClient` to throw an error.

- Use the `await` operator Asynchronously get the `PiiEntityCollection` for the run's InnerText.

- If there is PII in the text remove the `Text` child elements from r and append a new `Text` element with the redacted text.

- Optionally write data from each PII entity to the console.

```csharp
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
```

## 7. Save the new document with a new name

- At the end of the using statement use this code to get a new file name and save the file

```csharp
if (newDoc is not null)
{
    string[] strings = filePath.Split(".");
    string ext = strings[strings.Length - 1];
    strings = strings.SkipLast(1).ToArray();
    string root = String.Join(String.Empty, strings);

    newDoc.SaveAs($@"{root}_redacted.{ext}");
    newDoc.Close();
}
```

## 8. Sample Code

The following is the completed code. You can also view the complete [Program.cs here](../Program.cs).

```csharp
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
```
