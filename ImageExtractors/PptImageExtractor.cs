using DocSharp.Binary.PptFileFormat;
using DocSharp.Binary.StructuredStorage.Reader;
using Syncfusion.Drawing;



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
            string imageMediaType = Constants.MediaTypeConstants.MediaTypeUnknown;
            byte[]? imageBytes = null;
            if (pptDoc.PicturesContainer != null)
            {
                foreach (var pictureRecord in pptDoc.PicturesContainer._pictures.Values)
                {

                    
                    if (pictureRecord.RawData != null)
                    {

                        

                        //Code that appears to work to get image bytes, however images are coming not valid
                        //Create a stream from the raw data bytes
                        //using (var imageBytesStream = new MemoryStream(pictureRecord.RawData))
                        //{
                        //    imageBytes = imageBytesStream.ToArray();
                        //    imageBytesStream.Position = 0; // Ensure stream is at beginning before creating Image
                        //    var img = Image.FromStream(imageBytesStream);

                        //        imageMediaType = MimeTypeDetector.GetImageMediaType(imageBytes);
                        //        images.Add(new ExtractedImageInfo
                        //        {
                        //            ImageBytes = img.ImageData,
                        //            ImageMediaType = imageMediaType
                        //        });


                        //}

                        //Code that appears to work to get image bytes, however images are coming not valid
                        //using (var ms = new MemoryStream())
                        //{
                        //    pictureRecord.DumpToStream(ms);
                        //    imageBytes = ms.ToArray();
                        //}

                        //run the image bytes through the MediaTypeDetector to determine the actual image type
                        //if (imageBytes != null)
                        //{
                        //   imageMediaType = MimeTypeDetector.GetImageMediaType(imageBytes);
                        //   images.Add(new ExtractedImageInfo
                        //    {
                        //        ImageBytes = imageBytes,
                        //        ImageMediaType = imageMediaType
                        //    });
                        //}

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
                        // Commented out conversion: save original bytes and extension
                        // var conversion = ImageFormatConverter.ConvertToPng(image.ImageBytes, image.ImageMediaType);
                        string extension = ImageExtensionHelper.GetImageExtension(image.ImageMediaType);
                        byte[] bytesToWrite = image.ImageBytes;

                        imageCount++;
                        string fileName = Path.GetFileNameWithoutExtension(pptFile);
                        string imagePath = Path.Combine(pptOutputDir, $"{fileName}_Image_{imageCount}{extension}");
                        File.WriteAllBytes(imagePath, bytesToWrite);
                        Console.WriteLine($"    Saved image {imageCount}: {Path.GetFileName(imagePath)}");
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
