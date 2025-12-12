using SkiaSharp;

namespace LegacyPowerPointGetImages;

public class ConvertedImageResult
{
    public string ConversionStatus { get; set; } = string.Empty;
    public byte[]? ImageBytes { get; set; }
    public string? ErrorMessage { get; set; }
    public string OriginalMediaType { get; set; } = string.Empty;
    public bool ConversionRequired { get; set; }
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

            // If already a supported format (PNG, JPEG, WebP, or BMP), return as-is without conversion
            if (Array.Exists(SupportedMediaTypes, t => t.Equals(detectedMediaType, StringComparison.OrdinalIgnoreCase)))
            {
                return new ConvertedImageResult
                {
                    ConversionStatus = "Not Required",
                    ImageBytes = imageBytes,
                    OriginalMediaType = detectedMediaType,
                    ConversionRequired = false
                };
            }

            // For all other formats, conversion is required
            // Handle GIF conversion
            if (detectedMediaType == Constants.MediaTypeConstants.Gif)
            {
                using var bitmap = SKBitmap.Decode(imageBytes);
                if (bitmap == null)
                {
                    return new ConvertedImageResult
                    {
                        ConversionStatus = "Failed",
                        ErrorMessage = "Failed to decode GIF image with SkiaSharp.",
                        OriginalMediaType = detectedMediaType,
                        ConversionRequired = true
                    };
                }
                using var image = SKImage.FromBitmap(bitmap);
                using var data = image.Encode(SKEncodedImageFormat.Png, 100);
                if (data == null)
                {
                    return new ConvertedImageResult
                    {
                        ConversionStatus = "Failed",
                        ErrorMessage = "Failed to encode GIF as PNG with SkiaSharp.",
                        OriginalMediaType = detectedMediaType,
                        ConversionRequired = true
                    };
                }
                return new ConvertedImageResult
                {
                    ConversionStatus = "Success",
                    ImageBytes = data.ToArray(),
                    OriginalMediaType = detectedMediaType,
                    ConversionRequired = true
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
                        ConversionStatus = "Failed",
                        ErrorMessage = "Failed to decode TIFF image with SkiaSharp.",
                        OriginalMediaType = detectedMediaType,
                        ConversionRequired = true
                    };
                }
                using var image = SKImage.FromBitmap(bitmap);
                using var data = image.Encode(SKEncodedImageFormat.Png, 100);
                if (data == null)
                {
                    return new ConvertedImageResult
                    {
                        ConversionStatus = "Failed",
                        ErrorMessage = "Failed to encode TIFF as PNG with SkiaSharp.",
                        OriginalMediaType = detectedMediaType,
                        ConversionRequired = true
                    };
                }
                return new ConvertedImageResult
                {
                    ConversionStatus = "Success",
                    ImageBytes = data.ToArray(),
                    OriginalMediaType = detectedMediaType,
                    ConversionRequired = true
                };
            }

            // Attempt conversion for other types (ICO, etc.)
            using var genericBitmap = SKBitmap.Decode(imageBytes);
            if (genericBitmap == null)
            {
                return new ConvertedImageResult
                {
                    ConversionStatus = "Failed",
                    ErrorMessage = $"Failed to decode {detectedMediaType} image with SkiaSharp.",
                    OriginalMediaType = detectedMediaType,
                    ConversionRequired = true
                };
            }
            using var genericImage = SKImage.FromBitmap(genericBitmap);
            using var genericData = genericImage.Encode(SKEncodedImageFormat.Png, 100);
            if (genericData == null)
            {
                return new ConvertedImageResult
                {
                    ConversionStatus = "Failed",
                    ErrorMessage = $"Failed to encode {detectedMediaType} as PNG with SkiaSharp.",
                    OriginalMediaType = detectedMediaType,
                    ConversionRequired = true
                };
            }
            return new ConvertedImageResult
            {
                ConversionStatus = "Success",
                ImageBytes = genericData.ToArray(),
                OriginalMediaType = detectedMediaType,
                ConversionRequired = true
            };
        }
        catch (Exception ex)
        {
            return new ConvertedImageResult
            {
                ConversionStatus = "Failed",
                ErrorMessage = $"Image conversion error: {ex.Message}",
                OriginalMediaType = mediaType,
                ConversionRequired = true
            };
        }
    }
}
