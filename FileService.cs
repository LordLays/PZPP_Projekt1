using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PZPP_Projekt1
{
    static class FileService
    {
        public static string ReadFile(string filePath)
        {
            string content = string.Empty;

            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
                {
                    var paragraphs = doc.MainDocumentPart.RootElement.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        content += paragraph.InnerText + Environment.NewLine;
                    }
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Błąd odczytu pliku: {ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Wystąpił błąd: {ex.Message}", "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return content;
        }
    }
}
