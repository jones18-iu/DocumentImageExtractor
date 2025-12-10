namespace LegacyPowerPointGetImages;

public static class ImageExtensionHelper
{
    public static string GetImageExtension(string mediaType)
    {
        return mediaType switch
        {
            Constants.MediaTypeConstants.Jpeg => ".jpg",
            Constants.MediaTypeConstants.Png => ".png",
            Constants.MediaTypeConstants.Gif => ".gif",
            Constants.MediaTypeConstants.Bmp => ".bmp",
            Constants.MediaTypeConstants.Tiff => ".tiff",
            Constants.MediaTypeConstants.Webp => ".webp",
            _ => ".bin"
        };
    }
}
