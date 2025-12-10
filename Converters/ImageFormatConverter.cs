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
            // If already PNG, validate and return
            if (ValidatePngBytes(imageBytes))
            {
                return new ConvertedImageResult
                {
                    Success = true,
                    ImageBytes = imageBytes
                };
            }

            byte[] pngBytes;

            // Convert only if not supported
            if (!Array.Exists(SupportedMediaTypes, t => t.Equals(mediaType, StringComparison.OrdinalIgnoreCase)))
            {
                using var inputStream = new MemoryStream(imageBytes);
                using var image = Image.FromStream(inputStream);
                using var outputStream = new MemoryStream();
                image.Save(outputStream, ImageFormat.Png);
                pngBytes = outputStream.ToArray();
            }
            else
            {
                using var inStream = new MemoryStream(imageBytes);
                using var img = Image.FromStream(inStream);
                using var outStream = new MemoryStream();
                img.Save(outStream, ImageFormat.Png);
                pngBytes = outStream.ToArray();
            }

            if (!ValidatePngBytes(pngBytes))
            {
                return new ConvertedImageResult
                {
                    Success = false,
                    ErrorMessage = "Image conversion to PNG failed validation."
                };
            }

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

    private static bool ValidatePngBytes(byte[] bytes)
    {
        return bytes.Length > 7 &&
            bytes[0] == 0x89 &&
            bytes[1] == 0x50 &&
            bytes[2] == 0x4E &&
            bytes[3] == 0x47 &&
            bytes[4] == 0x0D &&
            bytes[5] == 0x0A &&
            bytes[6] == 0x1A &&
            bytes[7] == 0x0A;
    }
}
