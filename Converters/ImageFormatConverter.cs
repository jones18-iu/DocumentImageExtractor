using SkiaSharp;

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
        Constants.MediaTypeConstants.Webp,
        Constants.MediaTypeConstants.Bmp
    };

    public static ConvertedImageResult ConvertToPng(byte[] imageBytes, string mediaType)
    {
        try
        {
            // Detect the actual image type
            string detectedMediaType = MimeTypeDetector.GetImageMediaType(imageBytes);

            // If already a supported format (PNG, JPEG, WebP, or BMP), return as-is without conversion
            if (Array.Exists(SupportedMediaTypes, t => t.Equals(detectedMediaType, StringComparison.OrdinalIgnoreCase)))
            {
                return new ConvertedImageResult
                {
                    Success = true,
                    ImageBytes = imageBytes,
                    Extension = GetDefaultExtension(detectedMediaType)
                };
            }

            // Handle GIF conversion
            if (detectedMediaType == Constants.MediaTypeConstants.Gif)
            {
                using var bitmap = SKBitmap.Decode(imageBytes);
                if (bitmap == null)
                {
                    return new ConvertedImageResult
                    {
                        Success = false,
                        ErrorMessage = "Failed to decode GIF image with SkiaSharp."
                    };
                }
                using var image = SKImage.FromBitmap(bitmap);
                using var data = image.Encode(SKEncodedImageFormat.Png, 100);
                if (data == null)
                {
                    return new ConvertedImageResult
                    {
                        Success = false,
                        ErrorMessage = "Failed to encode GIF as PNG with SkiaSharp."
                    };
                }
                return new ConvertedImageResult
                {
                    Success = true,
                    ImageBytes = data.ToArray(),
                    Extension = ".png"
                };
            }

            // Handle TIFF conversion
            if (detectedMediaType == Constants.MediaTypeConstants.Tiff)
            {
                using var bitmap = SKBitmap.Decode(imageBytes);
                if (bitmap == null)
                {
                    return new ConvertedImageResult
                    {
                        Success = false,
                        ErrorMessage = "Failed to decode TIFF image with SkiaSharp."
                    };
                }
                using var image = SKImage.FromBitmap(bitmap);
                using var data = image.Encode(SKEncodedImageFormat.Png, 100);
                if (data == null)
                {
                    return new ConvertedImageResult
                    {
                        Success = false,
                        ErrorMessage = "Failed to encode TIFF as PNG with SkiaSharp."
                    };
                }
                return new ConvertedImageResult
                {
                    Success = true,
                    ImageBytes = data.ToArray(),
                    Extension = ".png"
                };
            }

            // Attempt conversion for other types (ICO, etc.)
            using var genericBitmap = SKBitmap.Decode(imageBytes);
            if (genericBitmap == null)
            {
                return new ConvertedImageResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to decode {detectedMediaType} image with SkiaSharp."
                };
            }
            using var genericImage = SKImage.FromBitmap(genericBitmap);
            using var genericData = genericImage.Encode(SKEncodedImageFormat.Png, 100);
            if (genericData == null)
            {
                return new ConvertedImageResult
                {
                    Success = false,
                    ErrorMessage = $"Failed to encode {detectedMediaType} as PNG with SkiaSharp."
                };
            }
            return new ConvertedImageResult
            {
                Success = true,
                ImageBytes = genericData.ToArray(),
                Extension = ".png"
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

    private static string GetDefaultExtension(string mediaType)
    {
        return mediaType switch
        {
            var t when t.Equals(Constants.MediaTypeConstants.Png, StringComparison.OrdinalIgnoreCase) => ".png",
            var t when t.Equals(Constants.MediaTypeConstants.Jpeg, StringComparison.OrdinalIgnoreCase) => ".jpg",
            var t when t.Equals(Constants.MediaTypeConstants.Webp, StringComparison.OrdinalIgnoreCase) => ".webp",
            var t when t.Equals(Constants.MediaTypeConstants.Bmp, StringComparison.OrdinalIgnoreCase) => ".bmp",
            _ => ".png"
        };
    }
}
