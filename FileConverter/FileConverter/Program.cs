namespace FileConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string basePath     = "C:/dev/FileConverter/FileConverter/FileConverter/tempFolder/";
            string wordPath     = $"{basePath}in.docx";
            string outputPath   = $"{basePath}out.pdf";

            WordToPdfConverter converter = new WordToPdfConverter();
            converter.ConvertWordToPdf(wordPath, outputPath);
        }
    }
}
