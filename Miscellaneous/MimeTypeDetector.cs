using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;


namespace LegacyPowerPointGetImages;


/// <summary>
/// Detects media (MIME) types from raw file bytes using signature heuristics.
/// Notes on accuracy and limitations:
/// - This is a pragmatic heuristic detector intended for typical PDF/DOC/DOCX/PPT/PPTX uploads.
/// - It does not perform spec-level parsing (e.g., ZIP central directory, OOXML content types, OLE directory entries).
/// - See detailed comments in methods for known edge cases and behaviors.
/// </summary>
public static class MimeTypeDetector
{
    /// <summary>
    /// Detects common image MIME types from byte signatures.
    /// Accuracy and limitations:
    /// - Matches JPEG/PNG/GIF/BMP/TIFF/WebP via well-known magic numbers.
    /// - Returns "**Unknown**" if unknown.
    /// - This is heuristic-only; uncommon or corrupted images may be misclassified or missed.
    /// </summary>
    /// <remarks>
    /// Expected byte signatures:
    /// JPEG: 255, 216, 255        (0xFF, 0xD8, 0xFF)
    /// PNG: 137, 80, 78, 71       (0x89, 0x50, 0x4E, 0x47)
    /// GIF: 71, 73, 70            (0x47, 0x49, 0x46)
    /// BMP: 66, 77                (0x42, 0x4D)
    /// TIFF (LE): 73, 73, 42, 0   (0x49, 0x49, 0x2A, 0x00)
    /// TIFF (BE): 77, 77, 0, 42   (0x4D, 0x4D, 0x00, 0x2A)
    /// WebP: 82, 73, 70, 70, ...  (0x52, 0x49, 0x46, 0x46, ... "WEBP" at offset 8)
    /// </remarks>
    /// <param name="bytes">Raw file bytes.</param>
    /// <returns>
    /// One of: "image/jpeg", "image/png", "image/gif", "image/bmp", "image/tiff", "image/webp",
    /// or "**Unknown**" if undetermined.
    /// </returns>
    public static string GetImageMediaType(byte[] bytes)
    {
        // JPEG: Starts with FF D8 FF
        if (bytes.Length > 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
        {
            return Constants.MediaTypeConstants.Jpeg;
        }

        // PNG: Starts with 89 50 4E 47
        if (bytes.Length > 8 && bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
        {
            return Constants.MediaTypeConstants.Png;
        }

        // GIF: Starts with 47 49 46 ("GIF")
        if (bytes.Length > 6 && bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46)
        {
            return Constants.MediaTypeConstants.Gif;
        }

        // BMP: Starts with 42 4D ("BM")
        if (bytes.Length > 1 && bytes[0] == 0x42 && bytes[1] == 0x4D)
        {
            return Constants.MediaTypeConstants.Bmp;
        }

        // TIFF: Starts with either 49 49 2A 00 (little endian) or 4D 4D 00 2A (big endian)
        if (bytes.Length > 3 &&
            ((bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2A && bytes[3] == 0x00) ||
             (bytes[0] == 0x4D && bytes[1] == 0x4D && bytes[2] == 0x00 && bytes[3] == 0x2A)))
        {
            return Constants.MediaTypeConstants.Tiff;
        }

        // WebP: Starts with "RIFF" (52 49 46 46) and "WEBP" (57 45 42 50) at offset 8
        if (bytes.Length > 11 &&
            bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46 && // "RIFF"
            bytes[8] == 0x57 && bytes[9] == 0x45 && bytes[10] == 0x42 && bytes[11] == 0x50) // "WEBP"
        {
            return Constants.MediaTypeConstants.Webp;
        }

        // Unknown type
        return Constants.MediaTypeConstants.MediaTypeUnknown;
      
    }

    /// <summary>
    /// Detects the MIME type of common document formats (PDF, legacy Office, OOXML) using byte signatures.
    /// 
    /// <para>Detection logic:</para>
    /// <list type="number">
    /// <item>PDF: Searches for "%PDF-" signature within first 2 MB (allows preamble bytes).</item>
    /// <item>Legacy Office (DOC/PPT): Detects CFBF header and scans for stream names.</item>
    /// <item>OOXML (DOCX/PPTX): Detects ZIP header and scans for Office directory names.</item>
    /// <item>Plain text: Heuristic check for printable ASCII.</item>
    /// </list>
    /// 
    /// <para>Returns:</para>
    /// <list type="bullet">
    /// <item>"application/pdf", "application/msword", "application/vnd.ms-powerpoint",</item>
    /// <item>"application/vnd.openxmlformats-officedocument.wordprocessingml.document",</item>
    /// <item>"application/vnd.openxmlformats-officedocument.presentationml.presentation",</item>
    /// <item>"text/plain", or "**Unknown**"</item>
    /// </list>
    /// <para>Limitations: Heuristic only; does not parse full file structures.</para>
    /// </summary>
    /// <param name="bytes">Raw file bytes.</param>
    /// <returns>Document MIME type or "**Unknown**".</returns>
    public static string GetDocumentMediaType(byte[] bytes)
    {
        // 1) PDF: "%PDF-" if these are present in file
        if (IsPdf(bytes))
        {
            return Constants.MediaTypeConstants.Pdf;
        }

        // 2) Legacy Office (DOC/PPT): Compound File Binary (CFBF) header: D0 CF 11 E0 A1 B1 1A E1
        // Improved: Scan for stream names in first 512 KB
        if (bytes.Length >= 8 &&
         bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0 &&
         bytes[4] == 0xA1 && bytes[5] == 0xB1 && bytes[6] == 0x1A && bytes[7] == 0xE1)
        {
            const string docStreamMarker = "WordDocument";
            const string pptStreamMarker = "PowerPoint Document";
            var scanLength = Math.Min(bytes.Length, 2 * 1024 * 1024);
            var asciiText = System.Text.Encoding.ASCII.GetString(bytes, 0, scanLength);
            var utf16leText = System.Text.Encoding.Unicode.GetString(bytes, 0, scanLength);

            if (asciiText.IndexOf(docStreamMarker, StringComparison.OrdinalIgnoreCase) >= 0 ||
                utf16leText.IndexOf(docStreamMarker, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return Constants.MediaTypeConstants.MsWord;
            }
            if (asciiText.IndexOf(pptStreamMarker, StringComparison.OrdinalIgnoreCase) >= 0 ||
                utf16leText.IndexOf(pptStreamMarker, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return Constants.MediaTypeConstants.MsPowerPoint;
            }
            return Constants.MediaTypeConstants.MediaTypeUnknown;
        }

        // 3) Office Open XML (DOCX/PPTX): ZIP local file header "PK\x03\x04"
        if (bytes.Length >= 4 && bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04)
        {
            var scanLength = Math.Min(bytes.Length, 2 * 1024 * 1024);
            var text = System.Text.Encoding.ASCII.GetString(bytes, 0, scanLength);
            if (text.IndexOf("word/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return Constants.MediaTypeConstants.MsWordX;
            }
            if (text.IndexOf("ppt/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return Constants.MediaTypeConstants.MsPowerPointX;
            }
            // Unknown type
            return Constants.MediaTypeConstants.MediaTypeUnknown;
        }

        // 4) Plain Text: Heuristic check for readable characters
        if (IsPlainText(bytes))
        {
            return Constants.MediaTypeConstants.PlainText;
        }

        // 5) Unknown type
        return Constants.MediaTypeConstants.MediaTypeUnknown;
    }



    /// <summary>
    /// Detects the MIME type of document formats using both byte signatures and file extension.
    /// 
    /// <para>Detection logic:</para>
    /// <list type="number">
    /// <item>PDF: Detected by bytes only ("%PDF-").</item>
    /// <item>Legacy Office/OOXML: Uses extension to confirm ambiguous types (e.g., .doc, .ppt, .docx, .pptx).</item>
    /// <item>Plain text: Detected by bytes only.</item>
    /// </list>
    /// 
    /// <para>Returns:</para>
    /// <list type="bullet">
    /// <item>"application/pdf", "application/msword", "application/vnd.openxmlformats-officedocument.wordprocessingml.document",</item>
    /// <item>"application/vnd.ms-powerpoint", "application/vnd.openxmlformats-officedocument.presentationml.presentation",</item>
    /// <item>"text/plain", or "**Unknown**"</item>
    /// </list>
    /// <para>Limitations: Uses extension for ambiguous Office types; does not parse full file structures.</para>
    /// </summary>
    /// <param name="bytes">Raw file bytes.</param>
    /// <param name="fileExtension">File extension - do not include the dot, e.g., "docx", "pdf").</param>
    /// <returns>Document MIME type or "**Unknown**".</returns>
    public static string GetDocumentMediaType(byte[] bytes, string fileExtension)
    {

        // 1) PDF: "%PDF-" if these are present in file
        if (IsPdf(bytes))
        {
            return Constants.MediaTypeConstants.Pdf;
        }

        // 2) Legacy Office (DOC/PPT): Compound File Binary (CFBF) header: D0 CF 11 E0 A1 B1 1A E1
        // Improved: Scan for stream names in first 512 KB
        if (bytes.Length >= 8 &&
         bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0 &&
         bytes[4] == 0xA1 && bytes[5] == 0xB1 && bytes[6] == 0x1A && bytes[7] == 0xE1)
        {
            if (fileExtension == "doc")
                return Constants.MediaTypeConstants.MsWord;
            if (fileExtension == "ppt")
                return Constants.MediaTypeConstants.MsPowerPoint;
            return Constants.MediaTypeConstants.MediaTypeUnknown;
        }

        // 3) Office Open XML (DOCX/PPTX): ZIP local file header "PK\x03\x04"
        if (bytes.Length >= 4 && bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04)
        {
            if (fileExtension == "docx")
                return Constants.MediaTypeConstants.MsWordX;
            if (fileExtension == "pptx")
                return Constants.MediaTypeConstants.MsPowerPointX;
            return Constants.MediaTypeConstants.MediaTypeUnknown;
        }

        // 4) Plain Text: Heuristic check for readable characters
        if (IsPlainText(bytes))
        {
            return Constants.MediaTypeConstants.PlainText;
        }

        // 5) Unknown type
        return Constants.MediaTypeConstants.MediaTypeUnknown;
    }

    /// <summary>
    /// Checks for the PDF signature ("%PDF-") within the first 2 MB of the file, allowing for preamble bytes.
    /// </summary>
    /// <param name="bytes">Raw file bytes.</param>
    /// <returns>True if PDF signature is found; otherwise false.</returns>
    private static bool IsPdf(byte[] bytes)
    {
        const byte percentByte = 0x25; // '%'
        const byte pByte = 0x50;       // 'P'
        const byte dByte = 0x44;       // 'D'
        const byte fByte = 0x46;       // 'F'
        const byte dashByte = 0x2D;    // '-'

        if (bytes.Length < 5)
            return false;

        // Limit search to 2 MB
        int searchLimit = Math.Min(bytes.Length, 2 * 1024 * 1024);

        for (int i = 0; i <= searchLimit - 5; i++)
        {
            if (bytes[i] == percentByte &&
                bytes[i + 1] == pByte &&
                bytes[i + 2] == dByte &&
                bytes[i + 3] == fByte &&
                bytes[i + 4] == dashByte)
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsPlainText(byte[] bytes)
    {
        const int MaxBytesToCheck = 1024; // Check the first 1 KB of the file
        int length = Math.Min(bytes.Length, MaxBytesToCheck);

        for (int i = 0; i < length; i++)
        {
            byte b = bytes[i];

            // Allow printable ASCII characters, newlines, tabs, and carriage returns
            if (b != '\n' && b != '\r' && b != '\t' && (b < 0x20 || b > 0x7E))
            {
                return false; // Non-printable character found
            }
        }

        return true; // All characters are printable
    }
}