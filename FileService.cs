using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextDocument = iTextSharp.text.Document;
using iTextPdfWriter = iTextSharp.text.pdf.PdfWriter;
using iTextParagraph = iTextSharp.text.Paragraph;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
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

        public static void SaveFile(string filePath, string content, bool asPdf = false)
        {
            if (asPdf)
            {
                try
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        iTextDocument doc = new iTextDocument();
                        iTextPdfWriter.GetInstance(doc, fs);
                        doc.Open();
                        doc.Add(new iTextParagraph(content));
                        doc.Close();
                    }
                    //MessageBox.Show("Plik zapisany jako PDF poprawnie.", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("Bląd przy zapisie pdf: " + ex.Message, "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                try
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart mainPart = doc.AddMainDocumentPart();
                        mainPart.Document = new Document();
                        Body body = mainPart.Document.AppendChild(new Body());
                        body.AppendChild(new Paragraph(new Run(new Text(content))));
                    }
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Bład przy zapisie pliku: " + ex.Message, "Błąd", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }


        }

    }
}
