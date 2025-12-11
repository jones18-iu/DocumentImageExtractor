using DocSharp.Binary.PptFileFormat;
using DocSharp.Binary.StructuredStorage.Reader;
using Syncfusion.Drawing;
using SkiaSharp;

namespace LegacyPowerPointGetImages;

public static class PptImageExtractor
{
    public static List<ExtractedImageInfo> ExtractImages(Stream pptStream)
    {
        var images = new List<ExtractedImageInfo>();
        try
        {

            throw new NotImplementedException("PPT image extraction does not work.");
            using var reader = new StructuredStorageReader(pptStream);
            var pptDoc = new PowerpointDocument(reader);

            if (pptDoc.PicturesContainer != null && pptDoc.PicturesContainer._pictures != null)
            {
                foreach (var pictureRecord in pptDoc.PicturesContainer._pictures.Values)
                {
                    if (pictureRecord.RawData == null || pictureRecord.RawData.Length == 0)
                    {
                        continue;
                    }

                    var rawBytes = pictureRecord.RawData;
                    var mediaType = MimeTypeDetector.GetImageMediaType(rawBytes);
                    var conversion = ImageFormatConverter.ConvertToPng(rawBytes, mediaType);
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
            return images;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting images from legacy PowerPoint document: {ex.Message}");
        }

        return images;
    }

    public static void ExtractAll(string pptInputDir, string pptOutputDir)
    {
        Console.WriteLine($"Processing PPT files from: {pptInputDir}");
        Console.WriteLine();

        if (!Directory.Exists(pptInputDir))
        {
            Console.WriteLine($"Error: Input directory not found at {pptInputDir}");
            return;
        }

        string[] pptFiles = Directory.GetFiles(pptInputDir, "*.ppt", SearchOption.TopDirectoryOnly);

        if (pptFiles.Length == 0)
        {
            Console.WriteLine("No PPT files found in the directory.");
            return;
        }

        Console.WriteLine($"Found {pptFiles.Length} PPT file(s) to process");
        Console.WriteLine();

        if (!Directory.Exists(pptOutputDir))
        {
            Directory.CreateDirectory(pptOutputDir);
            Console.WriteLine($"Created output directory: {pptOutputDir}");
        }

        foreach (string pptFile in pptFiles)
        {
            try
            {
                Console.WriteLine($"Processing: {Path.GetFileName(pptFile)}");
                using (FileStream fileStream = File.OpenRead(pptFile))
                {
                    var images = ExtractImages(fileStream);

                    if (images.Count == 0)
                    {
                        Console.WriteLine("  No images processed from this document.");
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
                        
                        string fileName = Path.GetFileNameWithoutExtension(pptFile);
                        string imagePath = Path.Combine(pptOutputDir, $"{fileName}_Image_{imageCount}{extension}");

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
                Console.WriteLine($"  Error processing {Path.GetFileName(pptFile)}: {ex.Message}");
                Console.WriteLine();
            }
        }

        Console.WriteLine($"Completed processing {pptFiles.Length} PPT file(s)");
    }
}
