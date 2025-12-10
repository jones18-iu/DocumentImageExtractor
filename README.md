# DocumentImageExtractor

DocumentImageExtractor extracts images from common document types and saves them to files or streams for further processing.

Supported formats
- PDF
- Microsoft Word (.docx) and legacy Word (.doc)
- Microsoft PowerPoint (.pptx)
- Legacy PowerPoint (.ppt) — work in progress / under development

Summary
This project provides tools and examples to locate and extract embedded images from documents (PDF, Word, PowerPoint). Extracted images can be saved to disk, copied to memory streams, uploaded, or passed to other processing pipelines.

Features
- Extracts images from PDF documents
- Extracts images from modern Word (.docx) documents
- Extracts images from modern PowerPoint (.pptx) presentations
- Provides support and guidance for legacy Word (.doc) files
- Legacy PowerPoint (.ppt) extraction is currently under development
- Saves images to files and exposes streams (MemoryStream) for custom handling
- Reports basic image metadata (MIME type, size)

Notes about legacy formats
- .docx and .pptx: These modern Office Open XML formats are typically handled using libraries such as DocumentFormat.OpenXml.
- Legacy .doc and .ppt: Older binary formats are not fully supported by the OpenXML libraries. To process legacy .doc or .ppt files you can:
  - Convert the file to .docx/.pptx using Microsoft Office (Save As) or another converter, then extract images from the converted file; or
  - Use Microsoft Office Interop (requires Microsoft Office installed) or a dedicated parser that supports the binary formats.
- This repository contains guidance and sample code for both modern formats and approaches for dealing with legacy files. Legacy PowerPoint extraction remains under active development — check the code and issues for progress.

Usage
1. Place sample documents in the TestData directory (or update the sample paths in code).
2. Update any file path variables in the sample program(s) if needed.
3. Run the application:
   ```bash
   dotnet run
   ```
4. The program will:
   - Open each document type you configure it to process
   - Enumerate embedded images
   - Report image metadata (content type and size)
   - Save each image as a file (e.g., Slide1_Image1.png, Document_Image1.jpg) or expose a stream for custom handling

Output example
```
Processing: TestData/Documents/sample.docx
  Found 4 images
    Image 1: image/png (45678 bytes) -> Saved as: Document_Image1.png
    Image 2: image/jpeg (23456 bytes) -> Saved as: Document_Image2.jpg

Processing: TestData/Documents/sample.pdf
  Found 2 images
    Image 1: image/png (12345 bytes) -> Saved as: PDF_Image1.png

Total images extracted: 6
```

Technical details
- Word/PowerPoint (.docx/.pptx): images are typically retrieved from package parts (ImagePart) and copied to MemoryStream for processing.
- PDF: images are extracted using a PDF parsing/extraction component that reads image XObjects and embeds them as image streams.
- Streams: extracted images are copied into MemoryStreams so they can be uploaded, transformed, or saved by downstream code.

Supported image MIME types
- image/png
- image/jpeg
- image/gif
- image/bmp
- image/tiff
- plus format-specific encodings found in documents

Requirements
- .NET 8.0 (or later runtime compatible with the project)
- DocumentFormat.OpenXml (for .docx/.pptx handling)
- A PDF extraction library (refer to project dependencies for the exact package used)
- Microsoft Office (optional) — only required if you choose to use Interop for legacy .doc/.ppt conversion or direct legacy-file processing

Project structure (example)
```
DocumentImageExtractor/
  src/
    Program.cs                # example entry point(s)
    Extractors/
      PdfImageExtractor.cs
      WordImageExtractor.cs
      PowerPointImageExtractor.cs
      LegacyWordHelper.cs
      LegacyPowerPointHelper.cs  # in-progress
  TestData/
    Documents/
      sample.docx
      sample.pdf
      sample.pptx
      sample_legacy.doc
      sample_legacy.ppt
  README.md
```

Custom processing
Instead of saving images to files you can:
- Upload to cloud storage
- Store in a database
- Send in an API response
- Process with image libraries (resize, convert, analyze)
- Any code that accepts a Stream can operate on the MemoryStream produced by the extractor

Error handling
The project includes handling for:
- Missing files and invalid paths
- Unsupported file formats
- Null or missing document parts
- Empty or corrupt documents (best-effort; behavior depends on parser used)

Contributing and development notes
- Contributions are welcome. If you work on legacy PowerPoint (.ppt) extraction, please open a draft PR or issue so progress can be tracked.
- For heavy legacy-format support, consider adding conversion steps to normalize to Open XML formats as part of a pipeline.

License
This sample project is provided for educational purposes. See LICENSE for details (if present).
