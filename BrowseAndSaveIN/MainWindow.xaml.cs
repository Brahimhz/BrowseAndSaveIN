using System;
using System.Windows;
using System.Windows.Forms;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

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
            var dialog = new FolderBrowserDialog();


            if (dialog.ShowDialog().Equals(System.Windows.Forms.DialogResult.OK))
                tbxFolder.Text = dialog.SelectedPath;

        }
    }
}
