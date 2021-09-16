using System;
using System.Collections.Generic;
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
        public MainWindow()
        {
            InitializeComponent();
        }

        private void buttonLoad_Click(object sender, RoutedEventArgs e)
        {
            Program program = new Program();

            OpenFileDialog dlg = new OpenFileDialog
            {
                Filter = "Excel документ (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
                CheckFileExists = true
            };

            if (dlg.ShowDialog() == true)
            {
                if (program.LoadXLSX(dlg.FileName).Length > 0)
                {
                    textBoxLoadedFile.Text = program.LoadXLSX(dlg.FileName);
                }
            }
        }

        private void buttonSave_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
