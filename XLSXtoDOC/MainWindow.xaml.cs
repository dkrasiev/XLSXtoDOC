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
        private string[,] _result;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void ButtonLoad_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel документ (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
                CheckFileExists = true
            };

            if (openFileDialog.ShowDialog() == true && openFileDialog.FileName != null)
            {
                _result = Program.LoadXLSX(openFileDialog.FileName);

                if (_result.Length > 0)
                {
                    textBlockSelectedFile.Text = openFileDialog.FileName;
                }
            }
        }

        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            if (_result == null)
            {
                Program.Error("Файл не выбран");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Word document (*.doc)|*.doc",
                CheckFileExists = false,
                CheckPathExists = true,
                FileName = "result.doc"
            };

            if (_result.Length < 1)
            {
                Program.Error("Файл не выбран или не содержит символов");
            }
            else if (saveFileDialog.ShowDialog() == true && saveFileDialog.FileName != null)
            {
                Program.SaveDOC(saveFileDialog.FileName, _result);
            }
        }
    }
}
