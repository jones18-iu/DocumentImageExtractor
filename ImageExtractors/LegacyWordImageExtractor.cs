using NPOI.HWPF;

namespace LegacyPowerPointGetImages;

public static class LegacyWordImageExtractor
{
    public static List<ExtractedImageInfo> ExtractImages(Stream wordStream)
    {
        var images = new List<ExtractedImageInfo>();
        try
        {
            if (!wordStream.CanSeek)
            {
                using var ms = new MemoryStream();
                wordStream.CopyTo(ms);
                wordStream = ms;
                wordStream.Position = 0;
            }
            var doc = new HWPFDocument(wordStream);
            var pictures = doc.GetPicturesTable().GetAllPictures();

            foreach (var pic in pictures)
            {
                var imageBytes = pic.GetContent();
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
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting images from legacy Word: {ex.Message}");
        }
        return images;
    }

    public static void ExtractAll(string docInputDir, string docOutputDir)
    {
        Console.WriteLine($"Processing DOC files from: {docInputDir}");
        Console.WriteLine();

        if (!Directory.Exists(docInputDir))
        {
            Console.WriteLine($"Error: Input directory not found at {docInputDir}");
            return;
        }

        string[] docFiles = Directory.GetFiles(docInputDir, "*.doc", SearchOption.TopDirectoryOnly);

        if (docFiles.Length == 0)
        {
            Console.WriteLine("No DOC files found in the directory.");
            return;
        }

        Console.WriteLine($"Found {docFiles.Length} DOC file(s) to process");
        Console.WriteLine();

        if (!Directory.Exists(docOutputDir))
        {
            Directory.CreateDirectory(docOutputDir);
            Console.WriteLine($"Created output directory: {docOutputDir}");
        }

        foreach (string docFile in docFiles)
        {
            try
            {
                Console.WriteLine($"Processing: {Path.GetFileName(docFile)}");
                using (FileStream fileStream = File.OpenRead(docFile))
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
                        
                        string fileName = Path.GetFileNameWithoutExtension(docFile);
                        string imagePath = Path.Combine(docOutputDir, $"{fileName}_Image_{imageCount}{extension}");

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
                Console.WriteLine($"  Error processing {Path.GetFileName(docFile)}: {ex.Message}");
                Console.WriteLine();
            }
        }

        Console.WriteLine($"Completed processing {docFiles.Length} DOC file(s)");
    }
}
