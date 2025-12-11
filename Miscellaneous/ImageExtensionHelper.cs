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
            Constants.MediaTypeConstants.Icon => ".ico",
            Constants.MediaTypeConstants.Svg => ".svg",
            Constants.MediaTypeConstants.Emf => ".emf",
            Constants.MediaTypeConstants.Wmf => ".wmf",
            Constants.MediaTypeConstants.Dds => ".dds",
            Constants.MediaTypeConstants.Jpeg2000 => ".jp2",
            _ => ".unknown"
        };
    }
}
