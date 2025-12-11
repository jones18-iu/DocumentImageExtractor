using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyPowerPointGetImages;


public static partial class Constants
{

    public static class MediaTypeConstants
    {
        //Document Media Types
        public const string Pdf = "application/pdf";
        public const string MsWord = "application/msword";
        public const string MsPowerPoint = "application/vnd.ms-powerpoint";
        public const string MsWordX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        public const string MsPowerPointX = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
        public const string PlainText = "text/plain";

        //Image Media Types
        public const string Jpeg = "image/jpeg";
        public const string Png = "image/png";
        public const string Gif = "image/gif";
        public const string Bmp = "image/bmp";
        public const string Tiff = "image/tiff";
        public const string Webp = "image/webp";
        public const string Icon = "image/x-icon";
        public const string Svg = "image/svg+xml";
        public const string Emf = "image/emf";
        public const string Wmf = "image/wmf";
        public const string Dds = "image/vnd.ms-dds";
        public const string Jpeg2000 = "image/jp2";

        //Media Type Not Found
        public const string MediaTypeUnknown = "**Unknown**";
    }
}