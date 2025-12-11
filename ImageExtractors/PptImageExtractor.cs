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
            using var reader = new StructuredStorageReader(pptStream);
            var pptDoc = new PowerpointDocument(reader);

            if (pptDoc.PicturesContainer != null && pptDoc.PicturesContainer._pictures != null)
            {
                foreach (var pictureRecord in pptDoc.PicturesContainer._pictures.Values)
                {
                    // Process only RawData and avoid DumpToStream
                    if (pictureRecord.RawData == null || pictureRecord.RawData.Length == 0)
                    {
                        continue;
                    }

                    var rawBytes = pictureRecord.RawData;

                    // Try SkiaSharp using encoded data helpers first
                    bool processed = false;
                    try
                    {
                        using var skImage = SKImage.FromEncodedData(rawBytes);
                        if (skImage != null)
                        {
                            using var skPng = skImage.Encode(SKEncodedImageFormat.Png, 100);
                            if (skPng != null)
                            {
                                images.Add(new ExtractedImageInfo
                                {
                                    ImageBytes = skPng.ToArray(),
                                    ImageMediaType = Constants.MediaTypeConstants.Png
                                });
                                processed = true;
                            }
                        }
                        if (!processed)
                        {
                            using var mem = new SKMemoryStream(rawBytes);
                            using var codec = SKCodec.Create(mem);
                            if (codec != null)
                            {
                                using var skBitmap = SKBitmap.Decode(codec);
                                if (skBitmap != null)
                                {
                                    using var skImage2 = SKImage.FromBitmap(skBitmap);
                                    using var skPng2 = skImage2.Encode(SKEncodedImageFormat.Png, 100);
                                    if (skPng2 != null)
                                    {
                                        images.Add(new ExtractedImageInfo
                                        {
                                            ImageBytes = skPng2.ToArray(),
                                            ImageMediaType = Constants.MediaTypeConstants.Png
                                        });
                                        processed = true;
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // ignore and fall back
                    }

                    if (!processed)
                    {
                        // Fall back to media type detection on RawData and return original bytes
                        var mediaType = MimeTypeDetector.GetImageMediaType(rawBytes);
                        images.Add(new ExtractedImageInfo
                        {
                            ImageBytes = rawBytes,
                            ImageMediaType = mediaType
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
                        // Use SkiaSharp via ImageFormatConverter to convert to PNG
                        var conversion = ImageFormatConverter.ConvertToPng(image.ImageBytes, image.ImageMediaType);
                        imageCount++;
                        string fileName = Path.GetFileNameWithoutExtension(pptFile);
                        string imagePath = Path.Combine(pptOutputDir, $"{fileName}_Image_{imageCount}{conversion.Extension}");

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
                Console.WriteLine($"  Error processing {Path.GetFileName(pptFile)}: {ex.Message}");
                Console.WriteLine();
            }
        }

        Console.WriteLine($"Completed processing {pptFiles.Length} PPT file(s)");
    }
}
