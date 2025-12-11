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
        string docOutputDir = Path.Combine(solutionRoot, "Output", "DocImages");
        string docxOutputDir = Path.Combine(solutionRoot, "Output", "DocxImages");
        string pptxOutputDir = Path.Combine(solutionRoot, "Output", "PptxImages");
        string pdfOutputDir = Path.Combine(solutionRoot, "Output", "PdfImages");
        string pptOutputDir = Path.Combine(solutionRoot, "Output", "PptImages");
        string pptxConvertedOutputDir = Path.Combine(solutionRoot, "Output", "PptxDocuments");

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
            Console.WriteLine("1. Extract images from all PDF files");
            Console.WriteLine("2. Extract images from all Word files (docx)");
            Console.WriteLine("3. Extract images from all PowerPoint files (pptx)");
            Console.WriteLine("4. Extract images from all legacy Word files  (doc)");
            Console.WriteLine("5. Extract images from all legacy PPT files (ppt) ");
            Console.WriteLine("6. Run All");
            Console.WriteLine("7. Convert legacy PPT files to PPTX (ppt -> pptx)");
            Console.WriteLine("0. Exit");
            Console.WriteLine();
            Console.Write("Enter your choice (0-7): ");

            string? choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    PdfImageExtractor.ExtractAll(pdfInputDir, pdfOutputDir);
                    break;
                case "2":
                    DocxImageExtractor.ExtractAll(docxInputDir, docxOutputDir);
                    break;
                case "3":
                    PptxImageExtractor.ExtractAll(pptxInputDir, pptxOutputDir);
                    break;
                case "4":
                    LegacyWordImageExtractor.ExtractAll(docInputDir, docOutputDir);
                    break;
                case "5":
                    PptImageExtractor.ExtractAll(pptInputDir, pptOutputDir);
                    break;
                case "6":
                    RunAll(pptInputDir, docInputDir, docOutputDir, docxInputDir, docxOutputDir, pptxInputDir, pptxOutputDir, pdfInputDir, pdfOutputDir, pptOutputDir);
                    break;
                case "7":
                    Console.WriteLine("Converting legacy PPT files to PPTX...");
                    PptToPptxConverter.ConvertAll(pptInputDir, pptxConvertedOutputDir);
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
    private static void RunAll(string pptInputDir, string docInputDir, string docOutputDir, string docxInputDir, string docxOutputDir, string pptxInputDir, string pptxOutputDir, string pdfInputDir, string pdfOutputDir, string pptOutputDir)
    {
        Console.WriteLine();
        Console.WriteLine("========================================");
        Console.WriteLine("  Running All Processing Tasks");
        Console.WriteLine("========================================");
        Console.WriteLine();

        Console.WriteLine("Task 1: Extracting images from all PDF files...");
        PdfImageExtractor.ExtractAll(pdfInputDir, pdfOutputDir);
        Console.WriteLine();

        Console.WriteLine("Task 2: Extracting images from all DOCX files...");
        DocxImageExtractor.ExtractAll(docxInputDir, docxOutputDir);
        Console.WriteLine();

        Console.WriteLine("Task 3: Extracting images from all PPTX files...");
        PptxImageExtractor.ExtractAll(pptxInputDir, pptxOutputDir);
        Console.WriteLine();

        Console.WriteLine("Task 4: Extracting images from all legacy Word (.doc) files...");
        LegacyWordImageExtractor.ExtractAll(docInputDir, docOutputDir);
        Console.WriteLine();

        Console.WriteLine("Task 5: Extracting images from all legacy PPT (.ppt) files...");
        PptImageExtractor.ExtractAll(pptInputDir, pptOutputDir);
        Console.WriteLine();
    }
    #endregion
}
