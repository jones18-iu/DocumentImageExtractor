using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyPowerPointGetImages;

public class ExtractedImageInfo
{
    public byte[]? ImageBytes { get; set; }
    public string ImageMediaType { get; set; } = string.Empty;
    public string ConversionStatus { get; set; } = string.Empty;
    public string? ConversionError { get; set; }
    public string OriginalMediaType { get; set; } = string.Empty;
    public bool ConversionRequired { get; set; }
}
