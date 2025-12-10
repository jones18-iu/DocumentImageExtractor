using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyPowerPointGetImages;

public class ExtractedImageInfo
{
    public required byte[] ImageBytes { get; set; }
    public string ImageMediaType { get; set; } = string.Empty;
}
