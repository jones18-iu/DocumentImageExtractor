using Syncfusion.Drawing;
using Syncfusion.Pdf.Parsing;

namespace LegacyPowerPointGetImages;

public static class PdfImageExtractor
{
    public static List<ExtractedImageInfo> ExtractImages(Stream pdfStream)
    {
        PdfDocumentExtractor documentExtractor = new PdfDocumentExtractor();
        var images = new List<ExtractedImageInfo>();
        try
        {
            documentExtractor.Load(pdfStream);

            Stream[] imageStreams = documentExtractor.ExtractImages();

            foreach (var stream in imageStreams)
            {
                var img = Image.FromStream(stream);
                var mediaType = MimeTypeDetector.GetImageMediaType(img.ImageData);
                var conversion = ImageFormatConverter.ConvertToPng(img.ImageData, mediaType);
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

            return images;
        }
        finally
        {
            documentExtractor.Dispose();
        }
    }

    public static void ExtractAll(string pdfInputDir, string pdfOutputDir)
    {
        Console.WriteLine($"Processing PDF files from: {pdfInputDir}");
        Console.WriteLine();

        if (!Directory.Exists(pdfInputDir))
        {
            Console.WriteLine($"Error: Input directory not found at {pdfInputDir}");
            return;
        }

        string[] pdfFiles = Directory.GetFiles(pdfInputDir, "*.pdf", SearchOption.TopDirectoryOnly);

        if (pdfFiles.Length == 0)
        {
            Console.WriteLine("No PDF files found in the directory.");
            return;
        }

        Console.WriteLine($"Found {pdfFiles.Length} PDF file(s) to process");
        Console.WriteLine();

        if (!Directory.Exists(pdfOutputDir))
        {
            Directory.CreateDirectory(pdfOutputDir);
            Console.WriteLine($"Created output directory: {pdfOutputDir}");
        }

        foreach (string pdfFile in pdfFiles)
        {
            try
            {
                Console.WriteLine($"Processing: {Path.GetFileName(pdfFile)}");
                using (FileStream fileStream = File.OpenRead(pdfFile))
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
                        
                        string fileName = Path.GetFileNameWithoutExtension(pdfFile);
                        string imagePath = Path.Combine(pdfOutputDir, $"{fileName}_Image_{imageCount}{extension}");

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
                Console.WriteLine($"  Error processing {Path.GetFileName(pdfFile)}: {ex.Message}");
                Console.WriteLine();
            }
        }

        Console.WriteLine($"Completed processing {pdfFiles.Length} PDF file(s)");
    }
}
