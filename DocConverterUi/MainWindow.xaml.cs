using DocConverter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DocConverterUi
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // TODO: Save in registry
        private string InputFileDirectory = @"C:\Users\dhfra\Documents\GitHub\DocConverter\TestFiles\";
        private string OutputFileDirectory = @"C:\Users\dhfra\Documents\GitHub\DocConverter\TestFiles\";
        private string DefaultOutputFilename = "";
        public MainWindow()
        {
            InitializeComponent();
            SetupDefaults();
        }

        private void SetupDefaults()
        {
            //throw new NotImplementedException();
        }

        private void InputFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            var openFileDlg = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".rtf",
                InitialDirectory = InputFileDirectory,
                //Multiselect = true,
                Title = "Input File"
            };


            // Launch OpenFileDialog by calling ShowDialog method
            var result = openFileDlg.ShowDialog();
            // Get the selected file name and display in a TextBox.
            // Load content of file in a TextBlock
            if (result == true)
            {
                var filename= openFileDlg.FileName;
                InputFileTextBox.Text = filename;
                InputFileDirectory = System.IO.Path.GetDirectoryName(filename);
                DefaultOutputFilename = System.IO.Path.GetFileNameWithoutExtension(filename);
                //TextBlock1.Text = System.IO.File.ReadAllText(openFileDlg.FileName);
            }
        }

        private void OutputFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            var openFileDlg = new Microsoft.Win32.SaveFileDialog
            {
                DefaultExt = ".docx",
                FileName = OutputFileDirectory + DefaultOutputFilename + ".docx",
                InitialDirectory = OutputFileDirectory,
                //Multiselect = true,
                Title = "Input File"
            };


            // Launch OpenFileDialog by calling ShowDialog method
            var result = openFileDlg.ShowDialog();
            // Get the selected file name and display in a TextBox.
            // Load content of file in a TextBlock
            if (result == true)
            {
                OuputFileTextBox.Text = openFileDlg.FileName;
                OutputFileDirectory = System.IO.Path.GetDirectoryName(openFileDlg.FileName);
                //TextBlock1.Text = System.IO.File.ReadAllText(openFileDlg.FileName);

                ConvertButton.IsEnabled = true;
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            var paragraphs = RtfReader.ReadParagraphs(InputFileTextBox.Text);

            var report = new Report();
            report.Parse(paragraphs);

            var docxOutputFile = OuputFileTextBox.Text;
            report.DocXWrite(docxOutputFile);

            ViewEditButton.IsEnabled = true;

        }

        private void ViewEditButton_Click(object sender, RoutedEventArgs e)
        {
            // @@@@@ TODO:  First check if file exists
            System.Diagnostics.Process.Start(OuputFileTextBox.Text);
        }
    }
}
