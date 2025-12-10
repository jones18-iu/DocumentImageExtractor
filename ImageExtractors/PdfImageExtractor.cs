using System;
using System.Collections.Generic;
using System.IO;
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
                using (var ms = new MemoryStream())
                {
                    stream.CopyTo(ms);
                    var imageBytes = ms.ToArray();
                    var mediaType = MimeTypeDetector.GetImageMediaType(imageBytes);
                    images.Add(new ExtractedImageInfo
                    {
                        ImageBytes = imageBytes,
                        ImageMediaType = mediaType
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
                        var conversion = ImageFormatConverter.ConvertToPng(image.ImageBytes, image.ImageMediaType);
                        string extension;
                        byte[] bytesToWrite;

                        if (conversion.Success && conversion.ImageBytes != null)
                        {
                            extension = conversion.Extension;
                            bytesToWrite = conversion.ImageBytes;
                        }
                        else
                        {
                            if (!conversion.Success)
                                Console.WriteLine($"    ! Conversion failed: {conversion.ErrorMessage}");
                            extension = ImageExtensionHelper.GetImageExtension(image.ImageMediaType);
                            bytesToWrite = image.ImageBytes;
                        }

                        imageCount++;
                        string fileName = Path.GetFileNameWithoutExtension(pdfFile);
                        string imagePath = Path.Combine(pdfOutputDir, $"{fileName}_Image_{imageCount}{extension}");
                        File.WriteAllBytes(imagePath, bytesToWrite);
                        Console.WriteLine($"    ? Saved image {imageCount}: {Path.GetFileName(imagePath)}");
                    }
                    Console.WriteLine();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ? Error processing {Path.GetFileName(pdfFile)}: {ex.Message}");
                Console.WriteLine();
            }
        }

        Console.WriteLine($"Completed processing {pdfFiles.Length} PDF file(s)");
    }
}
