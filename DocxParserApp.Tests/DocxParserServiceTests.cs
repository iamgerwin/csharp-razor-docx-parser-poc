using DocxParserApp.Services;
using Xunit;

namespace DocxParserApp.Tests;

public class DocxParserServiceTests
{
    private readonly DocxParserService _service;

    public DocxParserServiceTests()
    {
        _service = new DocxParserService();
    }

    [Fact]
    public void GenerateHtml_WithValidParseResult_ReturnsValidHtml()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            RawText = "Sample text",
            Headings = new List<HeadingElement>
            {
                new HeadingElement { Level = 1, Text = "Test Heading" }
            },
            Paragraphs = new List<ParagraphElement>
            {
                new ParagraphElement { Text = "Test paragraph", IsBold = false, IsItalic = false }
            }
        };

        // Act
        var html = _service.GenerateHtml(parseResult);

        // Assert
        Assert.Contains("<!DOCTYPE html>", html);
        Assert.Contains("<h1>", html);
        Assert.Contains("Test Heading", html);
        Assert.Contains("<p>", html);
        Assert.Contains("Test paragraph", html);
    }

    [Fact]
    public void GenerateHtml_WithBoldParagraph_IncludesBoldClass()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            Paragraphs = new List<ParagraphElement>
            {
                new ParagraphElement { Text = "Bold text", IsBold = true, IsItalic = false }
            }
        };

        // Act
        var html = _service.GenerateHtml(parseResult);

        // Assert
        Assert.Contains("class=\"bold\"", html);
        Assert.Contains("Bold text", html);
    }

    [Fact]
    public void GenerateHtml_WithItalicParagraph_IncludesItalicClass()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            Paragraphs = new List<ParagraphElement>
            {
                new ParagraphElement { Text = "Italic text", IsBold = false, IsItalic = true }
            }
        };

        // Act
        var html = _service.GenerateHtml(parseResult);

        // Assert
        Assert.Contains("class=\"italic\"", html);
        Assert.Contains("Italic text", html);
    }

    [Fact]
    public void GenerateHtml_WithTable_GeneratesTableHtml()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            Tables = new List<TableElement>
            {
                new TableElement
                {
                    Rows = new List<List<string>>
                    {
                        new List<string> { "Header 1", "Header 2" },
                        new List<string> { "Cell 1", "Cell 2" }
                    }
                }
            }
        };

        // Act
        var html = _service.GenerateHtml(parseResult);

        // Assert
        Assert.Contains("<table>", html);
        Assert.Contains("<th>", html);
        Assert.Contains("<td>", html);
        Assert.Contains("Header 1", html);
        Assert.Contains("Cell 1", html);
    }

    [Fact]
    public void GenerateHtml_EscapesSpecialCharacters()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            Paragraphs = new List<ParagraphElement>
            {
                new ParagraphElement { Text = "<script>alert('XSS')</script>", IsBold = false, IsItalic = false }
            }
        };

        // Act
        var html = _service.GenerateHtml(parseResult);

        // Assert
        Assert.DoesNotContain("<script>alert('XSS')</script>", html);
        Assert.Contains("&lt;script&gt;", html);
    }

    [Fact]
    public void GenerateJson_WithValidParseResult_ReturnsValidJson()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            RawText = "Sample text",
            Headings = new List<HeadingElement>
            {
                new HeadingElement { Level = 1, Text = "Test Heading" }
            },
            Paragraphs = new List<ParagraphElement>
            {
                new ParagraphElement { Text = "Test paragraph", IsBold = false, IsItalic = false }
            }
        };

        // Act
        var json = _service.GenerateJson(parseResult);

        // Assert
        Assert.Contains("\"rawText\"", json);
        Assert.Contains("\"headings\"", json);
        Assert.Contains("\"paragraphs\"", json);
        Assert.Contains("Sample text", json);
        Assert.Contains("Test Heading", json);
    }

    [Fact]
    public void GenerateJson_FormatsWithIndentation()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            RawText = "Test"
        };

        // Act
        var json = _service.GenerateJson(parseResult);

        // Assert
        Assert.Contains("\n", json); // Contains newlines (indented)
        Assert.Contains("  ", json); // Contains spaces (indented)
    }

    [Fact]
    public void GenerateJson_UsesCamelCase()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            RawText = "Test",
            Headings = new List<HeadingElement>
            {
                new HeadingElement { Level = 1, Text = "Heading" }
            }
        };

        // Act
        var json = _service.GenerateJson(parseResult);

        // Assert
        Assert.Contains("\"rawText\"", json); // camelCase
        Assert.Contains("\"headings\"", json); // camelCase
        Assert.DoesNotContain("\"RawText\"", json); // Not PascalCase
    }

    [Fact]
    public void HeadingElement_SetsPropertiesCorrectly()
    {
        // Arrange & Act
        var heading = new HeadingElement
        {
            Level = 2,
            Text = "Sample Heading"
        };

        // Assert
        Assert.Equal(2, heading.Level);
        Assert.Equal("Sample Heading", heading.Text);
    }

    [Fact]
    public void ParagraphElement_SetsPropertiesCorrectly()
    {
        // Arrange & Act
        var paragraph = new ParagraphElement
        {
            Text = "Sample paragraph",
            IsBold = true,
            IsItalic = true
        };

        // Assert
        Assert.Equal("Sample paragraph", paragraph.Text);
        Assert.True(paragraph.IsBold);
        Assert.True(paragraph.IsItalic);
    }

    [Fact]
    public void TableElement_InitializesWithEmptyRows()
    {
        // Arrange & Act
        var table = new TableElement();

        // Assert
        Assert.NotNull(table.Rows);
        Assert.Empty(table.Rows);
    }

    [Fact]
    public void DocxParseResult_InitializesCollections()
    {
        // Arrange & Act
        var result = new DocxParseResult();

        // Assert
        Assert.NotNull(result.Headings);
        Assert.NotNull(result.Paragraphs);
        Assert.NotNull(result.Tables);
        Assert.Empty(result.Headings);
        Assert.Empty(result.Paragraphs);
        Assert.Empty(result.Tables);
    }

    [Fact]
    public void GenerateHtml_WithMultipleHeadingLevels_GeneratesCorrectTags()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            Headings = new List<HeadingElement>
            {
                new HeadingElement { Level = 1, Text = "H1 Heading" },
                new HeadingElement { Level = 2, Text = "H2 Heading" },
                new HeadingElement { Level = 3, Text = "H3 Heading" }
            }
        };

        // Act
        var html = _service.GenerateHtml(parseResult);

        // Assert
        Assert.Contains("<h1>H1 Heading</h1>", html);
        Assert.Contains("<h2>H2 Heading</h2>", html);
        Assert.Contains("<h3>H3 Heading</h3>", html);
    }

    [Fact]
    public void GenerateJson_WithEmptyCollections_ReturnsValidJson()
    {
        // Arrange
        var parseResult = new DocxParseResult
        {
            RawText = string.Empty
        };

        // Act
        var json = _service.GenerateJson(parseResult);

        // Assert
        Assert.Contains("\"headings\": []", json);
        Assert.Contains("\"paragraphs\": []", json);
        Assert.Contains("\"tables\": []", json);
    }
}
