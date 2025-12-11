using System;
using System.IO;
using Syncfusion.Presentation;

namespace LegacyPowerPointGetImages;

public static class PptToPptxConverter
{
    /// <summary>
    /// Converts a single legacy PPT file to PPTX using Syncfusion Presentation.
    /// If the file cannot be opened/converted, an exception is thrown.
    /// </summary>
    /// <param name="pptFilePath">Input .ppt file path.</param>
    /// <param name="pptxOutputPath">Target .pptx file path.</param>
    public static void Convert(string pptFilePath, string pptxOutputPath)
    {
        if (string.IsNullOrWhiteSpace(pptFilePath))
            throw new ArgumentException("Input path is required", nameof(pptFilePath));
        if (string.IsNullOrWhiteSpace(pptxOutputPath))
            throw new ArgumentException("Output path is required", nameof(pptxOutputPath));
        if (!File.Exists(pptFilePath))
            throw new FileNotFoundException("Input PPT file not found", pptFilePath);

        Directory.CreateDirectory(Path.GetDirectoryName(pptxOutputPath)!);

        using FileStream input = File.OpenRead(pptFilePath);
        // Attempt to open using Syncfusion Presentation and save as PPTX
        using IPresentation presentation = Presentation.Open(input);
        using FileStream output = File.Create(pptxOutputPath);
        presentation.Save(output);
    }

    /// <summary>
    /// Converts all legacy PPT files in the input directory to PPTX files in the output directory using Syncfusion Presentation.
    /// </summary>
    /// <param name="pptInputDir">Directory containing .ppt files.</param>
    /// <param name="pptxOutputDir">Directory to write .pptx files.</param>
    public static void ConvertAll(string pptInputDir, string pptxOutputDir)
    {
        if (string.IsNullOrWhiteSpace(pptInputDir))
            throw new ArgumentException("Input directory is required", nameof(pptInputDir));
        if (string.IsNullOrWhiteSpace(pptxOutputDir))
            throw new ArgumentException("Output directory is required", nameof(pptxOutputDir));

        if (!Directory.Exists(pptInputDir))
        {
            Console.WriteLine($"Error: Input directory not found at {pptInputDir}");
            return;
        }

        Directory.CreateDirectory(pptxOutputDir);

        string[] pptFiles = Directory.GetFiles(pptInputDir, "*.ppt", SearchOption.TopDirectoryOnly);
        if (pptFiles.Length == 0)
        {
            Console.WriteLine("No PPT files found to convert.");
            return;
        }

        Console.WriteLine($"Found {pptFiles.Length} PPT file(s) to convert");
        foreach (var pptFile in pptFiles)
        {
            try
            {
                string nameNoExt = Path.GetFileNameWithoutExtension(pptFile);
                string targetPath = Path.Combine(pptxOutputDir, nameNoExt + ".pptx");
                Console.WriteLine($"Converting: {Path.GetFileName(pptFile)} -> {Path.GetFileName(targetPath)}");
                Convert(pptFile, targetPath);
                Console.WriteLine("  ✓ Converted successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  ✗ Failed to convert {Path.GetFileName(pptFile)}: {ex.Message}");
            }
        }

        Console.WriteLine("Completed PPT to PPTX conversion.");
    }
}
