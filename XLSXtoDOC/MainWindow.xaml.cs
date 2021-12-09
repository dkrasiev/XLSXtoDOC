using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using Microsoft.Win32;

namespace XLSXtoDOC
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string result = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonLoad_Click(object sender, RoutedEventArgs e)
        {
            Program program = new Program();

            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel документ (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
                CheckFileExists = true
            };

            if (dlg.ShowDialog() == true && dlg.FileName != null)
            {
                program.LoadXLSX(dlg.FileName, ref result);

                if (result.Length > 0)
                {
                    textBlockSelectedFile.Text = dlg.FileName;
                }
            }
        }

        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            Program program = new Program();

            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "Word document (*.doc)|*.doc",
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "result.doc"
            };

            if (result.Length < 1)
            {
                program.Error("Файл не выбран или не содержит символов");
            }
            else if (dlg.ShowDialog() == true && dlg.FileName != null)
            {
                program.SaveDOC(dlg.FileName, result);
            }
        }
    }
}
