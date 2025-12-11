using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace LegacyPowerPointGetImages;

public class ConvertedImageResult
{
    public bool Success { get; set; }
    public byte[]? ImageBytes { get; set; }
    public string Extension { get; set; } = ".png";
    public string? ErrorMessage { get; set; }
}

public static class ImageFormatConverter
{
    private static readonly string[] SupportedMediaTypes =
    {
        Constants.MediaTypeConstants.Png,
        Constants.MediaTypeConstants.Jpeg,
        Constants.MediaTypeConstants.Webp
    };

    public static ConvertedImageResult ConvertToPng(byte[] imageBytes, string mediaType)
    {
        try
        {
            // Detect the actual image type
            string detectedMediaType = MimeTypeDetector.GetImageMediaType(imageBytes);

            // If already a supported format (PNG, JPEG, or WebP), return as-is without conversion
            if (Array.Exists(SupportedMediaTypes, t => t.Equals(detectedMediaType, StringComparison.OrdinalIgnoreCase)))
            {
                return new ConvertedImageResult
                {
                    Success = true,
                    ImageBytes = imageBytes
                };
            }

            // Convert unsupported formats to PNG
            byte[] pngBytes;
            using var inputStream = new MemoryStream(imageBytes);
            using var image = Image.FromStream(inputStream);
            using var outputStream = new MemoryStream();
            image.Save(outputStream, ImageFormat.Png);
            pngBytes = outputStream.ToArray();

            return new ConvertedImageResult
            {
                Success = true,
                ImageBytes = pngBytes
            };
        }
        catch (Exception ex)
        {
            return new ConvertedImageResult
            {
                Success = false,
                ErrorMessage = $"Image conversion error: {ex.Message}"
            };
        }
    }
}
