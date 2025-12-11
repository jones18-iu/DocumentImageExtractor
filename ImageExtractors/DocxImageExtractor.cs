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
                            var conversion = ImageFormatConverter.ConvertToPng(imageBytes, mediaType);
                            if ((conversion.ConversionStatus == "Success" || conversion.ConversionStatus == "Not Required") && conversion.ImageBytes != null)
                            {
                                images.Add(new ExtractedImageInfo
                                {
                                    ImageBytes = conversion.ImageBytes,
                                    ImageMediaType = conversion.ConversionRequired ? Constants.MediaTypeConstants.Png : mediaType,
                                    ConversionStatus = conversion.ConversionStatus,
                                    OriginalMediaType = conversion.OriginalMediaType,
                                    ConversionRequired = conversion.ConversionRequired
                                });
                            }
                            else
                            {
                                images.Add(new ExtractedImageInfo
                                {
                                    ImageBytes = null,
                                    ImageMediaType = mediaType,
                                    ConversionStatus = conversion.ConversionStatus,
                                    ConversionError = conversion.ErrorMessage,
                                    OriginalMediaType = conversion.OriginalMediaType,
                                    ConversionRequired = conversion.ConversionRequired
                                });
                            }
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
                        imageCount++;
                        string extension = ImageExtensionHelper.GetImageExtension(image.ImageMediaType);
                        Console.WriteLine("  ----------------------------------------");
                        Console.WriteLine($"  Image {imageCount}:");
                        Console.WriteLine($"    Original Media Type: {image.OriginalMediaType}");
                        Console.WriteLine($"    Current Media Type: {image.ImageMediaType}");
                        Console.WriteLine($"    Extension: {extension}");
                        Console.WriteLine($"    Conversion Required: {image.ConversionRequired}");
                        Console.WriteLine($"    Conversion Status: {image.ConversionStatus}");
                        if (!string.IsNullOrEmpty(image.ConversionError))
                        {
                            Console.WriteLine($"    Conversion Error: {image.ConversionError}");
                        }
                        Console.WriteLine($"    Image Size: {(image.ImageBytes != null ? image.ImageBytes.Length : 0)} bytes");
                        
                        string fileName = Path.GetFileNameWithoutExtension(docxFile);
                        string imagePath = Path.Combine(docxOutputDir, $"{fileName}_Image_{imageCount}{extension}");

                        if ((image.ConversionStatus == "Success" || image.ConversionStatus == "Not Required") && image.ImageBytes != null)
                        {
                            File.WriteAllBytes(imagePath, image.ImageBytes);
                            Console.WriteLine($"    Saved: {Path.GetFileName(imagePath)}");
                        }
                        else
                        {
                            Console.WriteLine($"    Skipped: {(image.ImageBytes == null ? "Unsupported format" : "Conversion failed")}");
                        }
                        Console.WriteLine("  ----------------------------------------");
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
