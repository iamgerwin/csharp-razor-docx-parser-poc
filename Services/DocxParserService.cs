using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Text.Json;

namespace DocxParserApp.Services;

public class DocxParserService
{
    public DocxParseResult ParseDocument(string filePath)
    {
        var result = new DocxParseResult();

        using var wordDoc = WordprocessingDocument.Open(filePath, false);
        var body = wordDoc.MainDocumentPart?.Document?.Body;

        if (body == null)
        {
            result.RawText = "(Document body is empty)";
            return result;
        }

        var paragraphs = body.Elements<Paragraph>();
        var tables = body.Elements<Table>();
        var headings = new List<HeadingElement>();
        var paragraphList = new List<ParagraphElement>();
        var tableList = new List<TableElement>();
        var rawTextBuilder = new StringBuilder();

        // Parse paragraphs
        foreach (var para in paragraphs)
        {
            var text = para.InnerText;
            if (string.IsNullOrWhiteSpace(text)) continue;

            rawTextBuilder.AppendLine(text);

            // Check if it's a heading
            var pStyle = para.Elements<ParagraphProperties>()
                .FirstOrDefault()?
                .ParagraphStyleId?.Val?.Value;

            if (pStyle != null && pStyle.StartsWith("Heading"))
            {
                var level = int.TryParse(pStyle.Replace("Heading", ""), out var l) ? l : 1;
                headings.Add(new HeadingElement
                {
                    Level = level,
                    Text = text
                });
            }
            else
            {
                // Check for bold or italic
                var runs = para.Elements<Run>();
                var isBold = runs.Any(r => r.RunProperties?.Bold != null);
                var isItalic = runs.Any(r => r.RunProperties?.Italic != null);

                paragraphList.Add(new ParagraphElement
                {
                    Text = text,
                    IsBold = isBold,
                    IsItalic = isItalic
                });
            }
        }

        // Parse tables
        foreach (var table in tables)
        {
            var tableElement = new TableElement();
            foreach (var row in table.Elements<TableRow>())
            {
                var rowData = new List<string>();
                foreach (var cell in row.Elements<TableCell>())
                {
                    rowData.Add(cell.InnerText);
                }
                tableElement.Rows.Add(rowData);
            }
            tableList.Add(tableElement);
        }

        result.RawText = rawTextBuilder.ToString();
        result.Headings = headings;
        result.Paragraphs = paragraphList;
        result.Tables = tableList;

        return result;
    }

    public string GenerateHtml(DocxParseResult parseResult)
    {
        var html = new StringBuilder();
        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html>");
        html.AppendLine("<head>");
        html.AppendLine("<meta charset=\"utf-8\">");
        html.AppendLine("<style>");
        html.AppendLine("body { font-family: Arial, sans-serif; line-height: 1.6; padding: 20px; max-width: 800px; margin: 0 auto; }");
        html.AppendLine("h1, h2, h3, h4, h5, h6 { color: #333; margin-top: 1.5em; }");
        html.AppendLine("table { border-collapse: collapse; width: 100%; margin: 1em 0; }");
        html.AppendLine("td, th { border: 1px solid #ddd; padding: 8px; text-align: left; }");
        html.AppendLine("th { background-color: #f2f2f2; }");
        html.AppendLine(".bold { font-weight: bold; }");
        html.AppendLine(".italic { font-style: italic; }");
        html.AppendLine("</style>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");

        // Add headings and paragraphs in order
        foreach (var heading in parseResult.Headings)
        {
            html.AppendLine($"<h{heading.Level}>{System.Web.HttpUtility.HtmlEncode(heading.Text)}</h{heading.Level}>");
        }

        foreach (var para in parseResult.Paragraphs)
        {
            var cssClass = "";
            if (para.IsBold) cssClass += "bold ";
            if (para.IsItalic) cssClass += "italic";

            if (!string.IsNullOrEmpty(cssClass))
                html.AppendLine($"<p class=\"{cssClass.Trim()}\">{System.Web.HttpUtility.HtmlEncode(para.Text)}</p>");
            else
                html.AppendLine($"<p>{System.Web.HttpUtility.HtmlEncode(para.Text)}</p>");
        }

        // Add tables
        foreach (var table in parseResult.Tables)
        {
            html.AppendLine("<table>");
            bool firstRow = true;
            foreach (var row in table.Rows)
            {
                html.AppendLine("<tr>");
                var tag = firstRow ? "th" : "td";
                foreach (var cell in row)
                {
                    html.AppendLine($"<{tag}>{System.Web.HttpUtility.HtmlEncode(cell)}</{tag}>");
                }
                html.AppendLine("</tr>");
                firstRow = false;
            }
            html.AppendLine("</table>");
        }

        html.AppendLine("</body>");
        html.AppendLine("</html>");

        return html.ToString();
    }

    public string GenerateJson(DocxParseResult parseResult)
    {
        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        return JsonSerializer.Serialize(parseResult, options);
    }

    public string GenerateMarkdown(DocxParseResult parseResult)
    {
        var markdown = new StringBuilder();

        // Add headings
        foreach (var heading in parseResult.Headings)
        {
            var prefix = new string('#', heading.Level);
            markdown.AppendLine($"{prefix} {heading.Text}");
            markdown.AppendLine();
        }

        // Add paragraphs
        foreach (var para in parseResult.Paragraphs)
        {
            var text = para.Text;

            // Apply bold formatting
            if (para.IsBold && para.IsItalic)
            {
                text = $"***{text}***";
            }
            else if (para.IsBold)
            {
                text = $"**{text}**";
            }
            else if (para.IsItalic)
            {
                text = $"*{text}*";
            }

            markdown.AppendLine(text);
            markdown.AppendLine();
        }

        // Add tables
        foreach (var table in parseResult.Tables)
        {
            if (table.Rows.Count == 0) continue;

            // Header row
            if (table.Rows.Count > 0)
            {
                var headerRow = table.Rows[0];
                markdown.AppendLine("| " + string.Join(" | ", headerRow) + " |");

                // Separator row
                markdown.AppendLine("| " + string.Join(" | ", headerRow.Select(_ => "---")) + " |");

                // Data rows
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    markdown.AppendLine("| " + string.Join(" | ", table.Rows[i]) + " |");
                }

                markdown.AppendLine();
            }
        }

        return markdown.ToString().TrimEnd();
    }
}

public class DocxParseResult
{
    public string RawText { get; set; } = string.Empty;
    public List<HeadingElement> Headings { get; set; } = new();
    public List<ParagraphElement> Paragraphs { get; set; } = new();
    public List<TableElement> Tables { get; set; } = new();
}

public class HeadingElement
{
    public int Level { get; set; }
    public string Text { get; set; } = string.Empty;
}

public class ParagraphElement
{
    public string Text { get; set; } = string.Empty;
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
}

public class TableElement
{
    public List<List<string>> Rows { get; set; } = new();
}
