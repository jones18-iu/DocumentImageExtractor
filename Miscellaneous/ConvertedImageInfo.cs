namespace LegacyPowerPointGetImages;

public class ConvertedImageInfo
{
    public byte[] ImageBytes { get; }
    public string Extension => ".png";

    public ConvertedImageInfo(byte[] imageBytes)
    {
        ImageBytes = imageBytes;
    }
}