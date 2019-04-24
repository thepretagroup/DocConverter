using DocConverter;
using Microsoft.Win32;
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
        private string DefaultOutputFilename = "";

        private string InputFileDirectory
        {
            get => ReadFromRegistry(@"SOFTWARE\Preta", "InputFileDirectory", 
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            set => WriteToRegistry(@"SOFTWARE\Preta", "InputFileDirectory", value);
        }
        private string OutputFileDirectory {
            get => ReadFromRegistry(@"SOFTWARE\Preta", "OutputFileDirectory",
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            set => WriteToRegistry(@"SOFTWARE\Preta", "OutputFileDirectory", value);
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private string ReadFromRegistry(string subKey, string item, string defaultValue = "")
        {
            // open the subkey  
            RegistryKey key = Registry.CurrentUser.OpenSubKey(subKey);

            // if it does exist, retrieve the stored values  
            if (key == null)
            {
                return defaultValue;
            }
            else
            {
                var value = key.GetValue(item) as string ?? defaultValue;
                key.Close();
                return value;
            }
        }

        private void WriteToRegistry(string subKey, string item, string value)
        {
            // create the subkey  
            RegistryKey key = Registry.CurrentUser.CreateSubKey(subKey);
            key.SetValue(item, value);
            key.Close();
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
                var filename = openFileDlg.FileName; // First filename
                InputFileTextBox.Text = string.Join(Environment.NewLine, openFileDlg.FileNames);
                InputFileDirectory = System.IO.Path.GetDirectoryName(filename);
                DefaultOutputFilename = System.IO.Path.GetFileNameWithoutExtension(filename);

                OuputFileTextBox.Text = string.Empty;
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
                    InputFileTextBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None))
            {

                var paragraphs = RtfReader.ReadParagraphs(filename);
                var report = new Report();
                try
                {
                    report.Parse(paragraphs);
                }
                catch (ParseException ex)
                {
                    Console.WriteLine(">>> Input parsing error! {0}\n{1}", ex.Message, ex.InnerException.Message);

                    MessageBox.Show(ex.Message, "Invalid Input File", MessageBoxButton.OK,
                        MessageBoxImage.Exclamation);

                    return;
                }

                reports.Add(report);
            }

            if (reports.Count() > 0)
            {
                var docxOutputFile = OuputFileTextBox.Text;
                var reportWriter = new ReportWriter(reports);
                reportWriter.CreateDocX(docxOutputFile);

                ConvertButton.IsEnabled = false;
                ViewEditButton.IsEnabled = true;
            }
        }

        private void ViewEditButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(OuputFileTextBox.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Unable to Open '{0}'{1}{2}", OuputFileTextBox.Text,Environment.NewLine, ex.Message),
                    "Execution Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Exit application?", "Exit",
                MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) == MessageBoxResult.Yes)
            {
                Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;
                Application.Current.Shutdown();
            }
        }

        private void Image_MouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            MessageBox.Show(Properties.Resources.Instructions,"Instructions",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
