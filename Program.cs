using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System.Text;

namespace DocxToMarkdown;

public class Program
{
    public static async Task Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: DocxToMarkdown <input.docx> <output.md>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        try
        {
            await ConvertDocxToMarkdownAsync(inputPath, outputPath);
            Console.WriteLine($"Conversion complete: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }

    private static async Task ConvertDocxToMarkdownAsync(string inputPath, string outputPath)
    {
        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input DOCX file not found.", inputPath);
        }

        StringBuilder markdownBuilder = new();

        using WordprocessingDocument doc = WordprocessingDocument.Open(inputPath, false);
        MainDocumentPart? mainPart = doc.MainDocumentPart;
        if (mainPart != null)
        {
            // Process document structure
            ProcessDocumentStructure(mainPart, markdownBuilder);

            // Extract document metadata if needed
            ExtractMetadata(doc, markdownBuilder);
        }

        // Write markdown content to file
        await File.WriteAllTextAsync(outputPath, markdownBuilder.ToString());
    }

    private static void ProcessDocumentStructure(MainDocumentPart mainPart, StringBuilder markdownBuilder)
    {
        Body? body = mainPart.Document.Body;
        if (body == null)
        {
            return;
        }

        foreach (var element in body.Elements())
        {
            switch (element)
            {
                case Paragraph paragraph:
                    ProcessParagraph(paragraph, markdownBuilder);
                    break;
                case Table table:
                    ProcessTable(table, markdownBuilder);
                    break;
                case SectionProperties sectionProperties:
                    ProcessSectionProperties(sectionProperties, markdownBuilder);
                    break;
                default:
                    // Ignore other elements
                    break;
            }
        }
    }

    private static void ProcessParagraph(Paragraph paragraph, StringBuilder markdownBuilder)
    {
        // Determine paragraph style
        ParagraphProperties? paragraphProps = paragraph.ParagraphProperties;
        string? styleName = GetStyleName(paragraphProps);

        // Define the name of the custom style used for checklist items in Word
        const string ChecklistStyleName = "ChecklistItem"; // <<< BELANGRIJK: Moet overeenkomen met de naam in Word!

        // Handle heading levels
        if (styleName?.StartsWith("Heading") == true && int.TryParse(styleName.Substring(7), out int headingLevel))
        {
            string hashMarks = new('#', headingLevel);
            markdownBuilder.AppendLine($"{hashMarks} {ExtractTextFromParagraph(paragraph)}");
            return;
        }

        // --- NIEUWE CHECK: Handle Checklist items based on style ---
        if (styleName == ChecklistStyleName)
        {
            string checklistText = ExtractTextFromParagraph(paragraph).Trim();
            if (!string.IsNullOrWhiteSpace(checklistText))
            {
                markdownBuilder.AppendLine($"- [ ] {checklistText}");
                // GEEN extra AppendLine() hier, om "tight lists" te bevorderen
            }
            // Mogelijke lege checklist items negeren we
            return; // Checklist item verwerkt, ga niet verder
        }
        // --- EINDE NIEUWE CHECK ---

        // Handle bullet points
        if (HasBulletNumbering(paragraphProps))
        {
            markdownBuilder.AppendLine($"- {ExtractTextFromParagraph(paragraph)}");
            return;
        }

        // Regular paragraph
        string text = ExtractTextFromParagraph(paragraph);
        if (!string.IsNullOrWhiteSpace(text))
        {
            markdownBuilder.AppendLine($"{text}\n\n");
        }
    }

    private static string ExtractTextFromParagraph(Paragraph paragraph)
    {
        StringBuilder sb = new();

        foreach (var run in paragraph.Elements<Run>())
        {
            foreach (var textElement in run.Elements<Text>())
            {
                // Handle text formatting
                bool isBold = run.RunProperties?.Bold != null;
                bool isItalic = run.RunProperties?.Italic != null;

                string text = textElement.Text;

                if (isBold && isItalic)
                {
                    sb.Append($"***{text}***");
                }
                else if (isBold)
                {
                    sb.Append($"**{text}**");
                }
                else if (isItalic)
                {
                    sb.Append($"*{text}*");
                }
                else
                {
                    sb.Append(text);
                }
            }
        }

        return sb.ToString();
    }

    private static void ProcessTable(Table table, StringBuilder markdownBuilder)
    {
        // Process table rows
        List<TableRow> rows = table.Elements<TableRow>().ToList();
        if (!rows.Any())
        {
            return;
        }

        // Process header row
        TableRow headerRow = rows.First();
        List<TableCell> headerCells = headerRow.Elements<TableCell>().ToList();
        int columnCount = headerCells.Count;

        // Add header content
        markdownBuilder.Append("| ");
        foreach (var cell in headerCells)
        {
            markdownBuilder.Append(ExtractTextFromTableCell(cell));
            markdownBuilder.Append(" | ");
        }
        markdownBuilder.AppendLine();

        // Add separator row
        markdownBuilder.Append("| ");
        for (int i = 0; i < columnCount; i++)
        {
            markdownBuilder.Append("--- | ");
        }
        markdownBuilder.AppendLine();

        // Process data rows
        foreach (var row in rows.Skip(1))
        {
            markdownBuilder.Append("| ");
            foreach (var cell in row.Elements<TableCell>())
            {
                markdownBuilder.Append(ExtractTextFromTableCell(cell));
                markdownBuilder.Append(" | ");
            }
            markdownBuilder.AppendLine();
        }

        markdownBuilder.AppendLine();
    }

    private static string ExtractTextFromTableCell(TableCell cell)
    {
        StringBuilder sb = new();
        foreach (var paragraph in cell.Elements<Paragraph>())
        {
            sb.Append(ExtractTextFromParagraph(paragraph));
        }
        return sb.ToString().Replace("|", "\\|").Trim(); // Escape pipe characters and trim whitespace
    }

    private static void ProcessSectionProperties(SectionProperties sectionProperties, StringBuilder markdownBuilder)
    {
        // Add a horizontal line to separate sections
        markdownBuilder.AppendLine("\n---\n");
    }

    private static string? GetStyleName(ParagraphProperties? props)
    {
        if (props?.ParagraphStyleId?.Val != null)
        {
            return props.ParagraphStyleId.Val.Value;
        }
        return null;
    }

    private static bool HasBulletNumbering(ParagraphProperties? props)
    {
        return props?.NumberingProperties?.NumberingId != null;
    }

    private static void ExtractMetadata(WordprocessingDocument doc, StringBuilder markdownBuilder)
    {
        var coreProps = doc.PackageProperties;
        if (coreProps != null)
        {
            markdownBuilder.AppendLine("## Document Metadata");

            if (!string.IsNullOrEmpty(coreProps.Title))
            {
                markdownBuilder.AppendLine($"**Title**: {coreProps.Title}");
            }

            if (!string.IsNullOrEmpty(coreProps.Subject))
            {
                markdownBuilder.AppendLine($"**Subject**: {coreProps.Subject}");
            }

            if (!string.IsNullOrEmpty(coreProps.Creator))
            {
                markdownBuilder.AppendLine($"**Author**: {coreProps.Creator}");
            }

            if (coreProps.Created.HasValue)
            {
                markdownBuilder.AppendLine($"**Created**: {coreProps.Created.Value:yyyy-MM-dd HH:mm:ss}");
            }

            markdownBuilder.AppendLine();
        }
    }
}