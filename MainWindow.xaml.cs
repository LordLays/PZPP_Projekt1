using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System;
using System.Linq;


namespace PZPP_Projekt1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string currentFileName;

        public MainWindow()
        {
            InitializeComponent();
        }
      
        private void OpenButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Word Document (*.docx)|*.docx";
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                currentFileName = System.IO.Path.GetFileName(filePath);
                MainTextBox.Text = FileService.ReadFile(filePath);
                this.Title = currentFileName;
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter= "Word Document (*.docx)|*.docx|PDF files (*.pdf)|*.pdf";
            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;
                currentFileName = System.IO.Path.GetFileName(filePath);
                bool saveAsPdf = System.IO.Path.GetExtension(filePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase);
                FileService.SaveFile(filePath, MainTextBox.Text, saveAsPdf);
                this.Title = currentFileName;

            }
        }

       
    }
}