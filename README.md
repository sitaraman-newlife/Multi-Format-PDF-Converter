# Multi-Format PDF Converter

A powerful desktop utility for converting between PDF, Word (DOCX), Excel (XLSX), and image formats (JPEG/PNG) with batch processing capabilities, PDF merge/split functionality, and watermarking features.

## Features

### Core Conversion Features
- **PDF ↔ Word (DOCX)**: Convert PDF files to editable Word documents and vice versa
- **PDF ↔ Excel (XLSX)**: Extract tables from PDFs to Excel or create PDFs from spreadsheets
- **PDF ↔ Images**: Convert PDFs to images (JPEG/PNG) or create PDFs from image files
- **Batch Processing**: Convert multiple files simultaneously with progress tracking

### Advanced PDF Operations
- **Merge PDFs**: Combine multiple PDF files into a single document
- **Split PDFs**: Extract specific page ranges from PDF files
- **Watermark**: Add text or image watermarks to PDF documents

### User-Friendly Features
- Clean, intuitive Windows Forms interface
- Drag-and-drop file support
- Progress indicators for long-running operations
- Comprehensive error handling and logging
- File preview capabilities

## Technical Stack

- **.NET 8.0** - Modern cross-platform framework
- **Windows Forms** - Native Windows UI
- **iTextSharp 5.5.13.3** - PDF manipulation and generation
- **DocumentFormat.OpenXml 3.0.0** - Word/Excel file handling
- **PdfSharp 6.0.0** - PDF operations (merge, split)
- **System.Drawing.Common 8.0.0** - Image processing

## Prerequisites

- Windows 10 or later (64-bit)
- .NET 8.0 Runtime or SDK
- Visual Studio 2022 or later (for development)

## Installation & Setup

### For End Users

1. Download the latest release from the [Releases](../../releases) page
2. Extract the ZIP file to your desired location
3. Run `PdfConverter.exe`

### For Developers

1. Clone the repository:
```bash
git clone https://github.com/sitaraman-newlife/Multi-Format-PDF-Converter.git
cd Multi-Format-PDF-Converter
```

2. Open the solution in Visual Studio 2022:
```bash
cd PdfConverter
start PdfConverter.sln
```

3. Restore NuGet packages:
```bash
dotnet restore
```

4. Build the project:
```bash
dotnet build --configuration Release
```

5. Run the application:
```bash
dotnet run --project PdfConverter
```

## Project Structure

```
Multi-Format-PDF-Converter/
│
├── PdfConverter/                 # Main application project
│   ├── PdfConverter.csproj      # Project file with dependencies
│   ├── Program.cs               # Application entry point
│   ├── Forms/                   # Windows Forms UI
│   │   └── MainForm.cs         # Main application window
│   ├── Converters/              # Conversion logic
│   │   ├── PdfToWordConverter.cs
│   │   ├── WordToPdfConverter.cs
│   │   ├── PdfToExcelConverter.cs
│   │   ├── ExcelToPdfConverter.cs
│   │   ├── PdfToImageConverter.cs
│   │   └── ImageToPdfConverter.cs
│   ├── Operations/              # PDF operations
│   │   ├── PdfMerger.cs        # Merge multiple PDFs
│   │   ├── PdfSplitter.cs      # Split PDF by page range
│   │   └── WatermarkManager.cs # Add watermarks
│   └── Utils/                   # Utility classes
│       ├── FileValidator.cs
│       ├── Logger.cs
│       └── ProgressTracker.cs
│
├── Tests/                       # Unit tests
│   └── PdfConverter.Tests/
│
├── Docs/                        # Documentation
│   ├── UserGuide.md
│   └── DeveloperGuide.md
│
├── README.md                    # This file
├── LICENSE                      # MIT License
└── .gitignore                  # Git ignore rules
```

## Usage Guide

### Basic Conversion

1. Launch the application
2. Select the conversion type from the dropdown menu
3. Click "Add Files" or drag-and-drop files into the application
4. Configure conversion options (if applicable)
5. Click "Convert" to start the process
6. Converted files will be saved in the same directory or specified output folder

### Batch Processing

- Add multiple files using "Add Files" button or drag-and-drop
- All files will be processed sequentially
- Progress bar shows overall completion status
- Failed conversions are logged and reported

### Merge PDFs

1. Select "Merge PDFs" from the Operations menu
2. Add PDF files in desired order
3. Use up/down buttons to rearrange order
4. Click "Merge" and specify output filename

### Split PDFs

1. Select "Split PDF" from the Operations menu
2. Load a PDF file
3. Specify page range (e.g., "1-5, 10-15")
4. Click "Split" to extract pages

### Add Watermark

1. Select "Add Watermark" from the Operations menu
2. Choose text or image watermark
3. Configure position, opacity, and rotation
4. Apply to single PDF or batch process

## Building from Source

### Build Release Version

```bash
dotnet publish PdfConverter/PdfConverter.csproj -c Release -r win-x64 --self-contained
```

The compiled executable will be in `PdfConverter/bin/Release/net8.0-windows/win-x64/publish/`

### Run Tests

```bash
dotnet test
```

## Configuration

Application settings can be modified in `appsettings.json`:

```json
{
  "DefaultOutputFolder": "",
  "MaxBatchSize": 100,
  "LogLevel": "Information",
  "WatermarkDefaults": {
    "Opacity": 0.5,
    "Position": "Center"
  }
}
```

## Known Limitations

- PDF to Excel conversion works best with table-structured PDFs
- Complex PDF layouts may not convert perfectly to Word
- Image quality settings affect output file size
- Very large PDFs (100+ MB) may require additional processing time

## Troubleshooting

### "Failed to load PDF"
- Ensure PDF is not password-protected
- Check file is not corrupted
- Verify read permissions

### "Conversion failed"
- Check log file in `Logs/` folder for details
- Ensure sufficient disk space
- Verify input file format is supported

### Performance Issues
- Reduce batch size for large files
- Close other resource-intensive applications
- Consider processing files in smaller batches

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Libraries & References

- [iTextSharp](https://github.com/itext/itextsharp) - PDF manipulation
- [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) - Office document processing
- [PdfSharp](http://www.pdfsharp.net/) - PDF operations
- [.NET Documentation](https://docs.microsoft.com/dotnet/)

## Support

For issues, questions, or suggestions:
- Open an [Issue](../../issues)
- Check existing [Discussions](../../discussions)

## Author

**sitaraman-newlife**

## Acknowledgments

- iText team for iTextSharp library
- Microsoft for Open XML SDK
- PDFSharp team for PDF operations library
