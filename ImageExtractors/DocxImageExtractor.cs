using Syncfusion.DocIO.DLS;

namespace LegacyPowerPointGetImages;

public static class DocxImageExtractor
{
    public static List<ExtractedImageInfo> ExtractImages(Stream docxStream)
    {
        var images = new List<ExtractedImageInfo>();
        try
        {
            if (docxStream.CanSeek)
                docxStream.Position = 0;

            using var wordDocument = new WordDocument(docxStream, Syncfusion.DocIO.FormatType.Docx);

            foreach (WSection section in wordDocument.Sections)
            {
                foreach (WParagraph paragraph in section.Body.Paragraphs)
                {
                    foreach (Entity entity in paragraph.ChildEntities)
                    {
                        if (entity is WPicture picture && picture.ImageBytes != null)
                        {
                            var imageBytes = picture.ImageBytes;
                            var mediaType = MimeTypeDetector.GetImageMediaType(imageBytes);
                            images.Add(new ExtractedImageInfo
                            {
                                ImageBytes = imageBytes,
                                ImageMediaType = mediaType
                            });
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting images from DOCX: {ex.Message}");
        }
        return images;
    }

    public static void ExtractAll(string docxInputDir, string docxOutputDir)
    {
        Console.WriteLine($"Processing DOCX files from: {docxInputDir}");
        Console.WriteLine();

        if (!Directory.Exists(docxInputDir))
        {
            Console.WriteLine($"Error: Input directory not found at {docxInputDir}");
            return;
        }

        string[] docxFiles = Directory.GetFiles(docxInputDir, "*.docx", SearchOption.TopDirectoryOnly);

        if (docxFiles.Length == 0)
        {
            Console.WriteLine("No DOCX files found in the directory.");
            return;
        }

        Console.WriteLine($"Found {docxFiles.Length} DOCX file(s) to process");
        Console.WriteLine();

        if (!Directory.Exists(docxOutputDir))
        {
            Directory.CreateDirectory(docxOutputDir);
            Console.WriteLine($"Created output directory: {docxOutputDir}");
        }

        foreach (string docxFile in docxFiles)
        {
            try
            {
                Console.WriteLine($"Processing: {Path.GetFileName(docxFile)}");
                using (FileStream fileStream = File.OpenRead(docxFile))
                {
                    var images = ExtractImages(fileStream);

                    if (images.Count == 0)
                    {
                        Console.WriteLine("  No images found in this document.");
                        Console.WriteLine();
                        continue;
                    }

                    Console.WriteLine($"  Found {images.Count} image(s)");

                    int imageCount = 0;
                    foreach (var image in images)
                    {
                        var conversion = ImageFormatConverter.ConvertToPng(image.ImageBytes, image.ImageMediaType);
                        imageCount++;
                        string fileName = Path.GetFileNameWithoutExtension(docxFile);
                        string imagePath = Path.Combine(docxOutputDir, $"{fileName}_Image_{imageCount}{conversion.Extension}");

                        if (conversion.Success && conversion.ImageBytes != null)
                        {
                            File.WriteAllBytes(imagePath, conversion.ImageBytes);
                            Console.WriteLine($"    Saved converted image {imageCount}: {Path.GetFileName(imagePath)}");
                        }
                        else
                        {
                            Console.WriteLine($"    Conversion failed for image {imageCount}. Original image type: {image.ImageMediaType}. Skipping write of original.");
                        }
                    }
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  Error processing {Path.GetFileName(docxFile)}: {ex.Message}");
                Console.WriteLine();
            }
        }

        Console.WriteLine($"Completed processing {docxFiles.Length} DOCX file(s)");
    }
}
