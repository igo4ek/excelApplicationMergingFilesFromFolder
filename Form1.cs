using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace excelApplicationFindAndCopy
{
    public partial class Form1 : Form
    {
        string xlFileName = "";
        myExcelAdressMaker adrMaker = new myExcelAdressMaker();
        // приложение
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook workBook;

        // объект, содержащий все делегаты
        myDelegates delegates = new myDelegates();

        bool isDataGridViewFilled = false;

        public Form1()
        {
            InitializeComponent();
            // тут выставляем значения по умолчанию для элементов формы
            comboBox1.SelectedIndex = 1;
            comboBox2.SelectedIndex = 2;

            //ToolTip t6 = new ToolTip();
            //t6.SetToolTip(checkBox1, "Выводит в логи дополнительную информацию: что где найдено и какой диапазон ячеек при этом будет скопирован.\nНемного замедляет скорость отработки.");
        }



        // ЗАПОЛНЕНИЕ DATA_GRID_VIEW СОДЕРЖИМЫМ
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // Текущий лист для предосмотра
                Excel.Worksheet workSheet = workBook.Sheets[listBox1.SelectedIndex + 1];

                // Вызовем заполнение датаГридВью значениями из Текущего листа
                myExcel.setDataGridView(dataGridView1, workSheet, adrMaker, checkBox3);
            }
            catch (Exception exc)
            {
                MessageBox.Show("Во время предосмотра листа произошла ошибка.\n" + exc.Message + "\n\nОбратитесь к Разработчику.\nРекомендуется не сохранять изменения в файле.");
            }
            finally
            {
                isDataGridViewFilled = true;
            }
        }

        // для быстрого вывода текста на richTextBox2
        void print(string text)
        {
            richTextBox2.Invoke(delegates.richTextBoxUpdater, new object[] { text, richTextBox2 });
        }

        // Выбор папки и склейка всех находящихся так екселей в один
        private void button3_Click(object sender, EventArgs e)
        {
            int rowIgnoredBeforeList = comboBox1.SelectedIndex;
            int rowIgnoredAfterList = comboBox2.SelectedIndex;
            // Show the FolderBrowserDialog.
            DialogResult result = folderBrowserDialog1.ShowDialog();
            String path = "";
            if (result == DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;

                print("Шаг 1. Начало объединения файлов в один на отдельные листы.\n\n");
                string[] allFiles = Directory.GetFiles(path);
                // если файл всего один в папке
                if (allFiles.Length == 1)
                {
                    MessageBox.Show("Тут всего один файл.\nТут объединять нечего.\n\nВыберите папку, содержащую более чем один файл!");
                    return;
                }
                print("Порядок объединения файлов:\n");
                foreach (String file in allFiles)
                {
                    print(file + "\n");
                }
                // Объединение всех файлов в один
                for (int i = 1; i < allFiles.Length; i++)
                {
                    myExcel.copySheetFromFileToFile(allFiles[i], allFiles[0], "Лист" + (i + 1));
                }
                print("\nФайл из листов сформирован.\n\n");

                print("\nОткрытие файла...\n\n");
                // теперь все файлы находятся по пути allFiles[0] каждый файл в отдельном листе
                xlFileName = allFiles[0];
                if (!myExcel.checkWritingAvaliable(xlFileName, xlApp))
                {
                    print("!!! ФАЙЛ ЗАБЛОКИРОВАН ДРУГИМ ПОЛЬЗОВАТЕЛЕМ !!!\nЗакройте его в  Microsoft Excel и загрузите заново.\n\n");
                    xlFileName = "";
                    return;
                }


                // книга
                workBook = xlApp.Workbooks.Open(xlFileName, false, false); //открываем наш файл

                // количество листов:
                int sheetsCount = workBook.Worksheets.Count;

                print("Шаг 2. Начинаем объединять листы (" + sheetsCount + ")...\n");
                // Подготовка к объединению содержимого листов
                Excel.Worksheet pastedSheet = workBook.Worksheets[1];                       // лист, в который будем вставлять
                int rowPasted = myExcel.getRowsCount(pastedSheet) - rowIgnoredBeforeList - 1;   // строка, в которую будем вставлять: предпоследняя, потому что последняя не несёт важной информации
                xlApp.DisplayAlerts = false;

                Excel.Worksheet tmpCopySheet;   // лист откуда будем копировать
                int copyRowsCount;              // количество строк в листе, который будем копировать
                int copyColumnsCount;           // количество столбцов в листе, откуда будем копировать
                // начиная со второго листа будем копировать на первый
                print("Лист 1\n");
                for (int i = 2; i <= sheetsCount; i++)
                {
                    print("Лист " + i + "\n");

                    tmpCopySheet = workBook.Worksheets[i];
                    copyRowsCount = myExcel.getRowsCount(tmpCopySheet);
                    copyColumnsCount = myExcel.getColumnsCount(tmpCopySheet);
                    // 1. Выделение копируемого диапазона: начиная со (1+rowIgnoredBeforeList) строки по (copyRowsCount - rowIgnoredAfterList) строку
                    myExcel.CopyPasteRange(tmpCopySheet, pastedSheet, (1 + rowIgnoredBeforeList), 1, (copyRowsCount - rowIgnoredAfterList), copyColumnsCount, rowPasted, 1);
                    rowPasted += (copyRowsCount - (rowIgnoredAfterList + rowIgnoredBeforeList));
                }
                print("Листы объединены.\n\nПрименение автовысоты всем строкам...\n\n");
                // применим автовысоту для всех строк
                pastedSheet.Range[pastedSheet.Cells[1, 1], pastedSheet.Cells[myExcel.getRowsCount(pastedSheet), myExcel.getColumnsCount(pastedSheet)]].EntireRow.AutoFit(); //применить автовысоту;

                print("Удаление лишних листов...\n\n");
                // удалим все остальные листы, кроме первого
                for (int i = sheetsCount; i >= 2; i--)
                {
                    workBook.Worksheets[i].Delete();
                }
                print("Готово!");
                print(" ");

                //myExcel.setDataGridView(dataGridView1, pastedSheet, )
                listBox1.Invoke(delegates.listBoxUpdater, new object[] { listBox1, workBook });
                listBox1.Invoke(delegates.listBoxSelectedIndexChanger, new object[] { listBox1, 0 });
            }
            else
            {
                MessageBox.Show("Некорректно выбранная папка!", "Выбор папки...", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        // Действия при закрытии окна формы
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (xlFileName != "")
            {
                workBook.Close(true, xlFileName);
                xlApp.Quit();
            }
            System.Environment.Exit(1);
        }

        // автоскролл к введённой в textBox1 строке. по dataGridView1
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (isDataGridViewFilled) //!
            {
                int selectRow = 1;
                // нужно поределить номер строки, введённый в textBox1
                try { selectRow = Int32.Parse(textBox1.Text); }
                catch { return; }
                finally { }
                if (selectRow < 1 || selectRow > dataGridView1.Rows.Count) return;
                print("Распознано положительное число: " + selectRow + "\n");

                // снятие выделения
                dataGridView1.ClearSelection();
                // выделение нужной строки
                dataGridView1.Rows[selectRow - 1].Selected = true;
                // скролл к выделенной строке
                dataGridView1.FirstDisplayedScrollingRowIndex = selectRow - 1;
            }
        }

        // очистка textBox1
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }
    }
}
