using System.Text;

namespace LegacyPowerPointGetImages;

public static class ImageFormatDiagnoser
{
    public static string DiagnoseImageFormat(byte[] imageBytes)
    {
        if (imageBytes == null || imageBytes.Length < 4)
            return "Unknown (too small)";

        StringBuilder diagnosis = new StringBuilder();
        diagnosis.Append($"Hex start: {string.Join(" ", imageBytes.Take(16).Select(b => b.ToString("X2")))} | ");

        if (imageBytes.Length > 2 && imageBytes[0] == 0xFF && imageBytes[1] == 0xD8 && imageBytes[2] == 0xFF)
            diagnosis.Append("JPEG");
        else if (imageBytes.Length > 3 && imageBytes[0] == 0x89 && imageBytes[1] == 0x50 && imageBytes[2] == 0x4E && imageBytes[3] == 0x47)
            diagnosis.Append("PNG");
        else if (imageBytes.Length > 2 && imageBytes[0] == 0x47 && imageBytes[1] == 0x49 && imageBytes[2] == 0x46)
            diagnosis.Append("GIF");
        else if (imageBytes.Length > 1 && imageBytes[0] == 0x42 && imageBytes[1] == 0x4D)
            diagnosis.Append("BMP");
        else if (imageBytes.Length > 3 && ((imageBytes[0] == 0x49 && imageBytes[1] == 0x49 && imageBytes[2] == 0x2A && imageBytes[3] == 0x00) || (imageBytes[0] == 0x4D && imageBytes[1] == 0x4D && imageBytes[2] == 0x00 && imageBytes[3] == 0x2A)))
            diagnosis.Append("TIFF");
        else if (imageBytes.Length > 11 && imageBytes[0] == 0x52 && imageBytes[1] == 0x49 && imageBytes[2] == 0x46 && imageBytes[3] == 0x46 && imageBytes[8] == 0x57 && imageBytes[9] == 0x45 && imageBytes[10] == 0x42 && imageBytes[11] == 0x50)
            diagnosis.Append("WebP");
        else if (imageBytes.Length > 10 && imageBytes[0] == 0xFF && imageBytes[1] == 0xD8 && imageBytes[6] == 0x4A && imageBytes[7] == 0x46 && imageBytes[8] == 0x49 && imageBytes[9] == 0x46)
            diagnosis.Append("JPEG (JFIF)");
        else
            diagnosis.Append("UNKNOWN FORMAT");

        return diagnosis.ToString();
    }
}
