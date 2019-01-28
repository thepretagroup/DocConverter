using DocConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

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
        }

        private void InputFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            var openFileDlg = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".rtf",
                InitialDirectory = InputFileDirectory,
                Filter = "Rich Text Files|*.rtf",
                Title = "Input File"
            };

            // Launch OpenFileDialog by calling ShowDialog method
            var result = openFileDlg.ShowDialog();
            if (result == true && !string.IsNullOrEmpty(openFileDlg.FileName))
            {
                var filename= openFileDlg.FileName; // First filename
                InputFileTextBox.Text = string.Join(Environment.NewLine, openFileDlg.FileNames);
                InputFileDirectory = System.IO.Path.GetDirectoryName(filename);
                DefaultOutputFilename = System.IO.Path.GetFileNameWithoutExtension(filename);
                ConvertButton.IsEnabled = false;
                ViewEditButton.IsEnabled = false;
            }
        }

        private void OutputFileButton_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            var openFileDlg = new Microsoft.Win32.SaveFileDialog
            {
                DefaultExt = ".docx",
                FileName = DefaultOutputFilename + ".docx",
                InitialDirectory = OutputFileDirectory,
                Title = "Input File"
            };

            var result = openFileDlg.ShowDialog();
            if (result == true)
            {
                OuputFileTextBox.Text = openFileDlg.FileName;
                OutputFileDirectory = System.IO.Path.GetDirectoryName(openFileDlg.FileName);
                ViewEditButton.IsEnabled = false;
                ConvertButton.IsEnabled = true;
            }
        }

        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            var reports = new List<Report>();

            foreach (var filename in 
                    InputFileTextBox.Text.Split(new []{ Environment.NewLine }, StringSplitOptions.None)){

                var paragraphs = RtfReader.ReadParagraphs(filename);
                var report = new Report();
                report.Parse(paragraphs);
                reports.Add(report);
            }

            if (reports.Count() > 0)
            {
                var docxOutputFile = OuputFileTextBox.Text;
                var reportWriter = new ReportWriter(reports);
                reportWriter.CreateDocX(docxOutputFile);

                ViewEditButton.IsEnabled = true;
            }
        }

        private void ViewEditButton_Click(object sender, RoutedEventArgs e)
        {
            // @@@@@ TODO:  First check if file exists
            System.Diagnostics.Process.Start(OuputFileTextBox.Text);
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            if(MessageBox.Show("Exit application?", "Exit?", 
                MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;
                Application.Current.Shutdown();
            }
        }
    }
}
