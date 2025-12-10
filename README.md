# Legacy PowerPoint Image Extractor

This application extracts images from PowerPoint presentations and saves them to streams and files.

## Features

- Loads PowerPoint presentations
- Extracts all images from all slides
- Processes images as streams (MemoryStream)
- Saves extracted images to individual files
- Reports image metadata (content type, size)

## Important Note: .ppt vs .pptx Format

### Current Implementation
This program uses the **DocumentFormat.OpenXml** library, which **only supports .pptx** files (PowerPoint 2007 and later).

### Working with .ppt Files (Legacy Format)

The file `TestData\Documents\Image processing test ppt.ppt` is in the **legacy .ppt format**. To process it, you have two options:

#### Option 1: Convert to .pptx (Recommended)
1. Open the .ppt file in Microsoft PowerPoint
2. Click **File** ? **Save As**
3. Choose **PowerPoint Presentation (*.pptx)** as the file type
4. Save with a new name or replace the original

Then update the path in Program.cs:
```csharp
string pptFilePath = @"TestData\Documents\Image processing test ppt.pptx";
```

#### Option 2: Use Microsoft.Office.Interop.PowerPoint
For native .ppt support without conversion, you would need to:
1. Install Microsoft PowerPoint on the machine
2. Add `Microsoft.Office.Interop.PowerPoint` NuGet package
3. Use COM Interop to read the presentation

This approach is more complex and requires PowerPoint to be installed.

## Usage

1. Ensure your PowerPoint file is in .pptx format
2. Update the file path in `Program.cs` if needed:
   ```csharp
   string pptFilePath = @"TestData\Documents\Image processing test ppt.pptx";
   ```
3. Run the application:
   ```bash
   dotnet run
   ```

## Output

The program will:
- Display information about each slide processed
- Show the number of images found on each slide
- Display image metadata (content type, size in bytes)
- Save each image to a file named `Slide{N}_Image{M}.{ext}`
- Provide a total count of extracted images

Example output:
```
Loading PowerPoint file: TestData\Documents\Image processing test ppt.pptx

Found 3 slides

Processing Slide 1:
  Found 2 image(s)
    Image 1: image/png (45678 bytes)
      Saved as: Slide1_Image1.png
    Image 2: image/jpeg (23456 bytes)
      Saved as: Slide1_Image2.jpg

Processing Slide 2:
  No images found on this slide.

Processing Slide 3:
  Found 1 image(s)
    Image 1: image/png (34567 bytes)
      Saved as: Slide3_Image1.png

Total images extracted: 3
```

## Technical Details

### Image Processing
- Images are read from `ImagePart` objects using `GetStream()`
- Each stream is copied to a `MemoryStream` for processing
- Streams can be further processed, uploaded, or stored as needed

### Supported Image Formats
- PNG (image/png)
- JPEG (image/jpeg)
- GIF (image/gif)
- BMP (image/bmp)
- TIFF (image/tiff)
- EMF (image/x-emf)
- WMF (image/x-wmf)

## Requirements

- .NET 8.0
- DocumentFormat.OpenXml 3.0.2 (automatically restored)

## Project Structure

```
LegacyPowerPointGetImages/
??? Program.cs                          # Main application code
??? LegacyPowerPointGetImages.csproj   # Project file
??? TestData/
?   ??? Documents/
?       ??? Image processing test ppt.ppt  # Sample PowerPoint file
??? README.md                          # This file
```

## Code Modifications

To customize the image processing, modify the `ExtractImagesFromPresentation` method:

```csharp
// Instead of saving to file, you can:
// 1. Upload to cloud storage
// 2. Process with image libraries
// 3. Store in a database
// 4. Send over network
// 5. Apply transformations

using (Stream imageStream = imagePart.GetStream())
{
    using (MemoryStream memoryStream = new MemoryStream())
    {
        imageStream.CopyTo(memoryStream);
        
        // Your custom processing here
        // memoryStream contains the complete image data
    }
}
```

## Error Handling

The application includes error handling for:
- Missing files
- Unsupported file formats (.ppt vs .pptx)
- Null presentation parts
- Empty presentations

## License

This is a sample application for educational purposes.
