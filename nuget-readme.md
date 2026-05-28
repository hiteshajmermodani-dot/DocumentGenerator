# DocumentGeneratorLib

A powerful .NET library for creating and managing documents across multiple formats.

## Features

✨ **Multi-Format Support** - Generate PDF, and Word documents  
🎯 **Easy to Use** - Simple, intuitive API for document creation  
🔄 **Cross-Platform** - Works with .NET Framework 4.7.2 and .NET 9.0  
⚡ **High Performance** - Optimized for efficient document generation  

## Installation

Install DocxPDFGeneratorLib via NuGet Package Manager:

### .NET CLI
```bash
dotnet add package DocxPDFGeneratorLib
```

### Package Manager Console
```powershell
Install-Package DocxPDFGeneratorLib
```

### PackageReference
```xml
<PackageReference Include="DocxPDFGeneratorLib" Version="1.0.0" />
```

## Quick Start

### Creating a Word Document (.docx)

```csharp
using DocumentGenerator;

// Create a Word document with content
var doc = DocumentFactory.Create("output.docx");

doc.AddHeading("My Document Title", level: 1)
   .AddParagraph("This is a paragraph with some content.", fontSize: 11)
   .AddHeading("Section 1", level: 2)
   .AddParagraph("Here's some more text in section 1.", bold: true)
   .AddPageBreak()
   .AddHeading("Section 2", level: 2)
   .AddParagraph("Content in section 2.")
   .Save("output.docx");
```

### Creating a PDF Document (.pdf)

```csharp
using DocumentGenerator;

// Create a PDF document with content
var pdf = DocumentFactory.Create("output.pdf");

pdf.SetHeader("My Company")
   .SetFooter("Confidential", includePageNumbers: true)
   .AddHeading("Report Title", level: 1)
   .AddParagraph("This is the introduction paragraph.", fontSize: 11)
   .AddHeading("Overview", level: 2)
   .AddParagraph("Details about the report go here.")
   .Save("output.pdf");
```

## Documentation

For detailed documentation, examples, and API reference, please visit the [GitHub repository](https://github.com/hiteshajmermodani-dot/DocumentGenerator).

## Support

- 📧 **Issues**: Report bugs or request features on [GitHub Issues](https://github.com/hiteshajmermodani-dot/DocumentGenerator/issues)
- 💬 **Discussions**: Join the community on [GitHub Discussions](https://github.com/hiteshajmermodani-dot/DocumentGenerator/discussions)

## License

This project is licensed under the MIT License - see the [LICENSE](https://github.com/hiteshajmermodani-dot/DocumentGenerator/LICENSE.txt) file for details.

## Author

**Hitesh Modani**

---

⭐ If you find this library useful, please consider giving it a star on GitHub!
