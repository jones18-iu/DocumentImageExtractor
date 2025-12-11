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
                images.Add(new ExtractedImageInfo
                {
                    ImageBytes = imageBytes,
                    ImageMediaType = mediaType
                });
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
                        var conversion = ImageFormatConverter.ConvertToPng(image.ImageBytes, image.ImageMediaType);
                        imageCount++;
                        string fileName = Path.GetFileNameWithoutExtension(docFile);
                        string imagePath = Path.Combine(docOutputDir, $"{fileName}_Image_{imageCount}{conversion.Extension}");

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
                Console.WriteLine($"  Error processing {Path.GetFileName(docFile)}: {ex.Message}");
                Console.WriteLine();
            }
        }

        Console.WriteLine($"Completed processing {docFiles.Length} DOC file(s)");
    }
}
