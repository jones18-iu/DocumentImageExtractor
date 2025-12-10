using System.IO;
using System.Text;
using LegacyPowerPointGetImages;

namespace LegacyPowerPointGetImages;

class Program
{
   
    static void Main(string[] args)
    {
        #region Setup
        // Register encoding provider to support Windows-1252 and other legacy encodings
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Get the solution/project root directory
        string projectRoot = AppDomain.CurrentDomain.BaseDirectory;
        // Navigate up from bin/Debug/net8.0 to the project root
        DirectoryInfo? rootDir = new DirectoryInfo(projectRoot).Parent?.Parent?.Parent;
        string solutionRoot = rootDir?.FullName ?? AppDomain.CurrentDomain.BaseDirectory;

        // Input directory paths (relative to solution root)
        string pptInputDir = Path.Combine(solutionRoot, "TestData", "Documents", "Ppt");
        string docInputDir = Path.Combine(solutionRoot, "TestData", "Documents", "Doc");
        string docxInputDir = Path.Combine(solutionRoot, "TestData", "Documents", "Docx");
        string pptxInputDir = Path.Combine(solutionRoot, "TestData", "Documents", "Pptx");
        string pdfInputDir = Path.Combine(solutionRoot, "TestData", "Documents", "Pdf");

        // Output directory paths (relative to solution root)
        string pptConvertOutputDir = Path.Combine(solutionRoot, "Output", "PptImages");
        string docOutputDir = Path.Combine(solutionRoot, "Output", "DocImages");
        string docxOutputDir = Path.Combine(solutionRoot, "Output", "DocxImages");
        string pptxOutputDir = Path.Combine(solutionRoot, "Output", "PptxImages");
        string pdfOutputDir = Path.Combine(solutionRoot, "Output", "PdfImages");

        Console.WriteLine($"Solution Root: {solutionRoot}");
        Console.WriteLine();

    #endregion

    #region Menu
        while (true)
        {
            Console.Clear();
            Console.WriteLine("========================================");
            Console.WriteLine("  Document Image Extraction Tool");
            Console.WriteLine("========================================");
            Console.WriteLine();
            Console.WriteLine("Select an option:");
            Console.WriteLine("1. Convert all PPT files to PPTX");
            Console.WriteLine("2. Extract images from all legacy Word (.doc) files");
            Console.WriteLine("3. Extract images from all DOCX files");
            Console.WriteLine("4. Extract images from all PPTX files");
            Console.WriteLine("5. Extract images from all PDF files");
            Console.WriteLine("6. Run All");
            Console.WriteLine("7. Extract images from all legacy PPT (.ppt) files");
            Console.WriteLine("0. Exit");
            Console.WriteLine();
            Console.Write("Enter your choice (0-7): ");

            string? choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    PptToPptxConverter.ConvertAll(pptInputDir, pptConvertOutputDir);
                    break;
                case "2":
                    LegacyWordImageExtractor.ExtractAll(docInputDir, docOutputDir);
                    break;
                case "3":
                    DocxImageExtractor.ExtractAll(docxInputDir, docxOutputDir);
                    break;
                case "4":
                    PptxImageExtractor.ExtractAll(pptxInputDir, pptxOutputDir);
                    break;
                case "5":
                    PdfImageExtractor.ExtractAll(pdfInputDir, pdfOutputDir);
                    break;
                case "6":
                    RunAll(pptInputDir, pptConvertOutputDir, docInputDir, docOutputDir, docxInputDir, docxOutputDir, pptxInputDir, pptxOutputDir, pdfInputDir, pdfOutputDir);
                    break;
                case "7":
                    PptImageExtractor.ExtractAll();
                    break;
                case "0":
                    Console.WriteLine("Exiting...");
                    return;
                default:
                    Console.WriteLine("Invalid choice. Please try again.");
                    break;
            }

            if (choice != "0")
            {
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }
    #endregion
    }

    #region RunAll Method
    private static void RunAll(string pptInputDir, string pptConvertOutputDir, string docInputDir, string docOutputDir, string docxInputDir, string docxOutputDir, string pptxInputDir, string pptxOutputDir, string pdfInputDir, string pdfOutputDir)
    {
        Console.WriteLine();
        Console.WriteLine("========================================");
        Console.WriteLine("  Running All Processing Tasks");
        Console.WriteLine("========================================");
        Console.WriteLine();

        Console.WriteLine("Task 1: Converting all PPT files to PPTX...");
        PptToPptxConverter.ConvertAll(pptInputDir, pptConvertOutputDir);
        Console.WriteLine();

        Console.WriteLine("Task 2: Extracting images from all legacy Word (.doc) files...");
        LegacyWordImageExtractor.ExtractAll(docInputDir, docOutputDir);
        Console.WriteLine();

        Console.WriteLine("Task 3: Extracting images from all DOCX files...");
        DocxImageExtractor.ExtractAll(docxInputDir, docxOutputDir);
        Console.WriteLine();

        Console.WriteLine("Task 4: Extracting images from all PPTX files...");
        PptxImageExtractor.ExtractAll(pptxInputDir, pptxOutputDir);
        Console.WriteLine();

        Console.WriteLine("Task 5: Extracting images from all PDF files...");
        PdfImageExtractor.ExtractAll(pdfInputDir, pdfOutputDir);
        Console.WriteLine();

        Console.WriteLine("Task 6: Extracting images from all legacy PPT (.ppt) files...");
        PptImageExtractor.ExtractAll();
        Console.WriteLine();
    }
    #endregion
}
