using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
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

        public Form1()
        {
            InitializeComponent();
            // тут выставляем значения по умолчанию для элементов формы
           
            //ToolTip t6 = new ToolTip();
            //t6.SetToolTip(checkBox1, "Выводит в логи дополнительную информацию: что где найдено и какой диапазон ячеек при этом будет скопирован.\nНемного замедляет скорость отработки.");
        }

        // Кнопка "Файл..."
        
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

            }
        }

        // Выбор папки и склейка всех находящихся так екселей в один
        private void button3_Click(object sender, EventArgs e)
        {
            // Show the FolderBrowserDialog.
            DialogResult result = folderBrowserDialog1.ShowDialog();
            String path = "";
            if (result == DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;

                string[] allFiles = Directory.GetFiles(path);
                // Объединение всех файлов в один
                for (int i = 1; i < allFiles.Length; i++)
                {
                    myExcel.copySheetFromFileToFile(allFiles[i], allFiles[0], "Лист" + (i + 1));
                }

                richTextBox2.AppendText("Начало объединения файлов в один на отдельные листы\n");
                // теперь все файлы находятся по пути allFiles[0] каждый файл в отдельном листе
                xlFileName = allFiles[0];
                if (!myExcel.checkWritingAvaliable(xlFileName, xlApp))
                {
                    richTextBox2.AppendText("!!! ФАЙЛ ЗАБЛОКИРОВАН ДРУГИМ ПОЛЬЗОВАТЕЛЕМ !!!\nЗакройте его в  Microsoft Excel и загрузите заново.\n\n");
                    xlFileName = "";
                    return;
                }
                richTextBox2.AppendText("Файл из листов сформирован\n");

                // книга
                workBook = xlApp.Workbooks.Open(xlFileName, false, false); //открываем наш файл

                // количество листов:
                int sheetsCount = workBook.Worksheets.Count;

                richTextBox2.AppendText("Начинаем объединять листы...\n");
                richTextBox2.ScrollToCaret();
                // Подготовка к объединению содержимого листов
                Excel.Worksheet pastedSheet = workBook.Worksheets[1];   // лист, в который будем вставлять
                int rowPasted = myExcel.getRowsCount(pastedSheet) - 1;  // строка, в которую будем вставлять: предпоследняя, потому что последняя не несёт важной информации
                xlApp.DisplayAlerts = false;

                Excel.Worksheet tmpCopySheet;   // лист откуда будем копировать
                int copyRowsCount;              // количество строк в листе, который будем копировать
                int copyColumnsCount;           // количество столбцов в листе, откуда будем копировать
                // начиная со второго листа будем копировать на первый
                for (int i = 2; i <= sheetsCount; i++)
                {
                    richTextBox2.AppendText("Лист "+i+"\n");
                    richTextBox2.ScrollToCaret();

                    tmpCopySheet = workBook.Worksheets[i];
                    copyRowsCount = myExcel.getRowsCount(tmpCopySheet);
                    copyColumnsCount = myExcel.getColumnsCount(tmpCopySheet);
                    // 1. Выделение копируемого диапазона: начиная со 2 строки по copyRowsCount - 2 строку
                    myExcel.CopyPasteRange(tmpCopySheet, pastedSheet, 2, 1, copyRowsCount - 2, copyColumnsCount, rowPasted, 1);
                    rowPasted += (copyRowsCount - 3);
                    //if (i == 2) rowPasted--; // после первой вставки из-за заголовка, который нужно будет впредь игнорировать
                }
                richTextBox2.AppendText("Листы объединены.\n");
                richTextBox2.ScrollToCaret();
                // применим автовысоту для всех строк
                pastedSheet.Range[pastedSheet.Cells[1, 1], pastedSheet.Cells[myExcel.getRowsCount(pastedSheet), myExcel.getColumnsCount(pastedSheet)]].EntireRow.AutoFit(); //применить автовысоту;

                // удалим все остальные листы, кроме первого
                for (int i = sheetsCount; i >= 2; i--)
                {
                    workBook.Worksheets[i].Delete();
                }    
                MessageBox.Show("Файлы объединены");

                //myExcel.setDataGridView(dataGridView1, pastedSheet, )
                listBox1.Invoke(delegates.listBoxUpdater, new object[] { listBox1, workBook });
                listBox1.SelectedIndex = 0;
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
    }
}
