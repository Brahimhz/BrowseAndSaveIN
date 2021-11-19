using Microsoft.Win32;
using System;
using System.Windows;

namespace BrowseAndSaveIN
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog fileDialog = new OpenFileDialog();

            fileDialog.Multiselect = false;
            fileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            fileDialog.DefaultExt = ".xlsx";
            Nullable<bool> dialogOk = fileDialog.ShowDialog();

            if (dialogOk == true)
                tbxFiles.Text = fileDialog.FileName;

        }

        private void btnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.FolderBrowserDialog result = dialog.ShowDialog();
            tbxFolder.Text = dialog.SelectedPath;
        }
    }
}
