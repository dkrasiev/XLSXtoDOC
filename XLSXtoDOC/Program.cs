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
    static class Program
    {
        /// <summary>
        /// Получение данных из excel
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="result"></param>
        public static string[,] LoadXLSX(string filename)
        {
            string[,] result = null;

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
                    Excel.Range usedRange = worksheet.UsedRange;
                    // Получаем строки в используемом диапазоне
                    Excel.Range urRows = usedRange.Rows;
                    // Получаем столбцы в используемом диапазоне
                    Excel.Range urColums = usedRange.Columns;

                    // Количества строк и столбцов
                    int rowsCount = urRows.Count;
                    int columnsCount = urColums.Count;
                    result = new string[rowsCount, columnsCount];
                    for (int i = 1; i <= rowsCount; i++)
                    {
                        for (int j = 1; j <= columnsCount; j++)
                        {
                            Excel.Range cellRange = usedRange.Cells[i, j];
                            // Получение текста ячейки
                            string cellText = (cellRange == null || cellRange.Value2 == null) ? null :
                                                (cellRange as Excel.Range).Value2.ToString();

                            if (cellText != null)
                            {
                                result[i - 1, j - 1] += cellText;
                            }
                        }
                    }
                    // Очистка неуправляемых ресурсов на каждой итерации
                    if (urRows != null) Marshal.ReleaseComObject(urRows);
                    if (urColums != null) Marshal.ReleaseComObject(urColums);
                    if (usedRange != null) Marshal.ReleaseComObject(usedRange);
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

            return result;
        }

        /// <summary>
        /// Сохранение таблицы в word
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="result"></param>
        public static void SaveDOC(string filename, string[,] result)
        {
            object SaveChanges = false;

            Word.Application app = new Word.Application();
            Word.Document doc = app.Documents.Add();
            Word.Range range = doc.Range();

            // Определение количества необходимых строк с столбцов
            int rowCount = result.GetLength(0);
            int columnCount = result.GetLength(1);

            try
            {
                // Создание и добавление таблицы
                Word.Table table = doc.Tables.Add(range, rowCount, columnCount);

                // Заполнение таблицы данными
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        table.Cell(i, j).Range.Text = result[i, j];
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
        public static void Error(string message)
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
        public static void Message(string message)
        {
            string messageBoxText = message;
            string caption = "Message";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Information;

            MessageBox.Show(messageBoxText, caption, button, icon);
        }
    }
}
