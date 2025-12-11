using Syncfusion.Presentation;

namespace LegacyPowerPointGetImages;

public static class PptxImageExtractor
{
    public static List<ExtractedImageInfo> ExtractImages(Stream pptxStream)
    {
        var images = new List<ExtractedImageInfo>();
        try
        {
            if (pptxStream.CanSeek)
                pptxStream.Position = 0;

            using var presentation = Presentation.Open(pptxStream);

            foreach (ISlide slide in presentation.Slides)
            {
                foreach (IShape shape in slide.Shapes)
                {
                    if (shape is IPicture picture && picture.ImageData != null)
                    {
                        var imageBytes = picture.ImageData;
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
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting images from PPTX: {ex.Message}");
        }
        return images;
    }

    public static void ExtractAll(string pptxInputDir, string pptxOutputDir)
    {
        Console.WriteLine($"Processing PPTX files from: {pptxInputDir}");
        Console.WriteLine();

        if (!Directory.Exists(pptxInputDir))
        {
            Console.WriteLine($"Error: Input directory not found at {pptxInputDir}");
            return;
        }

        string[] pptxFiles = Directory.GetFiles(pptxInputDir, "*.pptx", SearchOption.TopDirectoryOnly);

        if (pptxFiles.Length == 0)
        {
            Console.WriteLine("No PPTX files found in the directory.");
            return;
        }

        Console.WriteLine($"Found {pptxFiles.Length} PPTX file(s) to process");
        Console.WriteLine();

        if (!Directory.Exists(pptxOutputDir))
        {
            Directory.CreateDirectory(pptxOutputDir);
            Console.WriteLine($"Created output directory: {pptxOutputDir}");
        }

        foreach (string pptxFile in pptxFiles)
        {
            try
            {
                Console.WriteLine($"Processing: {Path.GetFileName(pptxFile)}");
                using (FileStream fileStream = File.OpenRead(pptxFile))
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
                        string fileName = Path.GetFileNameWithoutExtension(pptxFile);
                        string imagePath = Path.Combine(pptxOutputDir, $"{fileName}_Image_{imageCount}{conversion.Extension}");

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
                Console.WriteLine($"  Error processing {Path.GetFileName(pptxFile)}: {ex.Message}");
                Console.WriteLine();
            }
        }

        Console.WriteLine($"Completed processing {pptxFiles.Length} PPTX file(s)");
    }
}
