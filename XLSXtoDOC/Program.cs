using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows;
using System.IO;

namespace XLSXtoDOC
{
    class Program
    {
        /// <summary>
        /// Получение данных из excel
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="result"></param>
        public void LoadXLSX(string filename, ref string result)
        {
            result = "";

            string FileName = filename;
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            Excel.Application app = new Excel.Application();
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            try
            {
                workbooks = app.Workbooks;
                workbook = workbooks.Open(FileName, MissingObj, rOnly);

                // Получение всех страниц докуента
                sheets = workbook.Sheets;

                foreach (Excel.Worksheet worksheet in sheets)
                {
                    // Получаем диапазон используемых на странице ячеек
                    Excel.Range UsedRange = worksheet.UsedRange;
                    // Получаем строки в используемом диапазоне
                    Excel.Range urRows = UsedRange.Rows;
                    // Получаем столбцы в используемом диапазоне
                    Excel.Range urColums = UsedRange.Columns;

                    // Количества строк и столбцов
                    int RowsCount = urRows.Count;
                    int ColumnsCount = urColums.Count;
                    for (int i = 1; i <= RowsCount; i++)
                    {
                        for (int j = 1; j <= ColumnsCount; j++)
                        {
                            Excel.Range CellRange = UsedRange.Cells[i, j];
                            // Получение текста ячейки
                            string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                                (CellRange as Excel.Range).Value2.ToString();

                            if (CellText != null)
                            {
                                result += CellText;
                            }

                            if (j != ColumnsCount)
                            {
                                result += "\t";
                            }
                            else if (i != RowsCount)
                            {
                                result += "\n";
                            }
                        }
                    }
                    // Очистка неуправляемых ресурсов на каждой итерации
                    if (urRows != null) Marshal.ReleaseComObject(urRows);
                    if (urColums != null) Marshal.ReleaseComObject(urColums);
                    if (UsedRange != null) Marshal.ReleaseComObject(UsedRange);
                    if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                }
            }
            catch (Exception ex)
            {
                /* Обработка исключений */
                Error(ex.Message);
            }
            finally
            {
                /* Очистка оставшихся неуправляемых ресурсов */
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (workbook != null)
                {
                    workbook.Close(SaveChanges);
                    Marshal.ReleaseComObject(workbook);
                    workbook = null;
                }

                if (workbooks != null)
                {
                    workbooks.Close();
                    Marshal.ReleaseComObject(workbooks);
                    workbooks = null;
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
            }
        }

        /// <summary>
        /// Сохранение таблицы в word
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="input"></param>
        public void SaveDOC(string filename, string input)
        {
            object SaveChanges = false;

            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            Word.Range range = doc.Range();

            // Определение количества необходимых строк с столбцов
            int rowCount = 1;
            int columnCount = 0;
            foreach (char s in input)
            {
                if (s == '\n')
                {
                    rowCount++;
                }
                if (s == '\t')
                {
                    columnCount++;
                }
            }
            columnCount = (columnCount + rowCount) / rowCount;

            // Создание массива для заполнения таблицы
            string[] content = input.Replace("\n", "\t").Split('\t');

            try
            {
                // Создание и добавление таблицы
                Word.Table table = doc.Tables.Add(range, rowCount, columnCount);

                // Заполнение таблицы данными
                int a = 1;
                int b = 1;
                for (int i = 0; i < rowCount * columnCount; i++)
                {
                    table.Cell(a, b).Range.Text = content[i];
                    if (b == columnCount)
                    {
                        b = 1;
                        a++;
                    }
                    else
                    {
                        b++;
                    }
                }

                // Сохранение документа
                doc.SaveAs(filename);
                Message("Файл успешно сохранен");
            }
            catch (Exception ex)
            {
                /* Обработка исключений */
                Error(ex.Message);
            }
            finally
            {
                /* Очистка оставшихся неуправляемых ресурсов */
                if (doc != null)
                {
                    doc.Close(ref SaveChanges);
                }
                if (range != null)
                {
                    Marshal.ReleaseComObject(range);
                    range = null;
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
            }
        }

        /// <summary>
        /// Вывод окна с сообщением об ошибке
        /// </summary>
        /// <param name="message"></param>
        public void Error(string message)
        {
            string messageBoxText = message;
            string caption = "Error";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Error;

            MessageBox.Show(messageBoxText, caption, button, icon);
        }

        /// <summary>
        /// Вывод окна с сообщением
        /// </summary>
        /// <param name="message"></param>
        public void Message(string message)
        {
            string messageBoxText = message;
            string caption = "Message";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Information;

            MessageBox.Show(messageBoxText, caption, button, icon);
        }
    }
}
