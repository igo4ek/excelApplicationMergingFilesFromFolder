using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Runtime.CompilerServices;
using System.Threading;

namespace excelApplicationFindAndCopy
{
    class myExcel
    {
        /// <summary>
        /// На вход получает Timespan (разницу между двумя DateTime.Now)
        /// На выходе строка с временем выполнения кода.
        /// </summary>
        /// <param name="time">= DateTime two DateTime.Now {код программы} минус другой DateTime one DateTime.Now</param>
        /// <returns></returns>
        public static string getTime(TimeSpan time)
        {
            string result = "0 мс";
            if (time.Milliseconds != 0)
            {
                result = time.Milliseconds + " мс";
            }
            if (time.Seconds != 0)
            {
                result = time.Seconds + " cек. " + time.Milliseconds + " мс";
            }
            if (time.Minutes != 0)
            {
                result = time.Minutes + " мин. " + time.Seconds + " cек. " + time.Milliseconds + " мс";
            }
            if (time.Hours != 0)
            {
                result = time.Hours + " часов. " + time.Minutes + " мин. " + time.Seconds + " cек. " + time.Milliseconds + " мс";
            }
            return result;
        }

        /// <summary>
        /// Возвращает количество строк на листе sheet
        /// </summary>
        /// <param name="sheet"> Лист, для которого нужно узнать количество строк</param>
        /// <returns></returns>
        public static int getRowsCount(Excel.Worksheet sheet)
        {
            try
            {
                return sheet.Cells.Find("*", System.Reflection.Missing.Value,
                                   System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                   Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                   false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            }
            catch
            {
                return sheet.Cells[sheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row; //последняя заполненная строка в столбце А
            }
            finally { }
        }

        /// <summary>
        /// Возвращает максимально часто встречающееся число строк
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="adrMaker"></param>
        /// <returns></returns>
        public static int getMaxMeetingRowsCount(Excel.Worksheet sheet, myExcelAdressMaker adrMaker)
        {
            // отберем количества строк в КАЖДОМ столбце
            int columnsCount = getColumnsCount(sheet);  // количество столбцов
            List<int> rowsCountList = new List<int>();  // количество строк в каждом столбце
            for (int j = 1; j <= columnsCount; j++)
            {
                rowsCountList.Add(sheet.Cells[sheet.Rows.Count, adrMaker.getLetter(j)].End[Excel.XlDirection.xlUp].Row); // последняя заполненная строка в столбце j
            }

            // отсортируем по возрастанию
            rowsCountList.Sort();

            // найдем наиболее часто встречающееся число количества столбцов
            int presentValue = rowsCountList[0];
            int count = 0;
            List<int> valuesList = new List<int>(); // список оригинальных значений
            List<int> countsList = new List<int>(); // список количеств этих значений
            valuesList.Add(presentValue);
            for (int i = 0; i < rowsCountList.Count; i++)
            {
                if (rowsCountList[i] != presentValue)
                {
                    valuesList.Add(rowsCountList[i]);
                    countsList.Add(count);

                    // подготовка к дальнейшему перебору
                    presentValue = rowsCountList[i];
                    count = 0;
                }
                count++;
                if (i == rowsCountList.Count - 1)
                {
                    countsList.Add(count);
                }
            }
            return valuesList[countsList.IndexOf(countsList.Max())]; // число, встречавшееся максимальное количество раз
        }


        /// <summary>
        /// Возвращает количество столбцов на листе sheet...
        /// </summary>
        /// <param name="sheet">Лист, для которого нужно узнать количество столбцов</param>
        /// <returns></returns>
        public static int getColumnsCount(Excel.Worksheet sheet)
        {
            try
            {
                return sheet.Columns.Cells.Find("*", System.Reflection.Missing.Value,
                                   System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                   Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                   false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
            }
            catch 
            { 
                return sheet.Cells[1, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column; //последний заполненный столбец в 1-й строке    
            }
            finally { }
        }

        /// <summary>
        /// Копирует и вставляет диапазон из одного листа в другой
        /// </summary>
        /// <param name="sheetCopy">Лист откуда будем копировать</param>
        /// <param name="sheetPaste">Лист куда будем вставлять</param>
        /// <param name="rowStart"></param>
        /// <param name="columnStart"></param>
        /// <param name="rowEnd"></param>
        /// <param name="columnEnd"></param>
        /// <param name="rowPasted"></param>
        /// <param name="columnPasted"></param>
        public static void CopyPasteRange(Excel.Worksheet sheetCopy, Excel.Worksheet sheetPaste, int rowStart, int columnStart, int rowEnd, int columnEnd, int rowPasted, int columnPasted)
        {
            sheetCopy.Select();
            sheetCopy.Range[sheetCopy.Cells[rowStart, columnStart], sheetCopy.Cells[rowEnd, columnEnd]].Copy();
            sheetPaste.Select();
            sheetPaste.Cells[rowPasted, columnPasted].PasteSpecial();
        }

        /// <summary>
        /// Закрашивает фон указанной ячейки в цвет
        /// </summary>
        /// <param name="sheet">Лист, на котором будем закрашивать фон ячейки</param>
        /// <param name="row">Строка</param>
        /// <param name="column">Столбец</param>
        /// <param name="color">Цвет</param>
        public static void fillOnColor(Excel.Worksheet sheet, int row, int column, Color color)
        {
            sheet.Cells[row, column].Interior.Color = color;
        }


        /// <summary>
        /// Проверяет не заблокирован ли файл другим пользователем
        /// </summary>
        /// <param name="xlFileName">Путь_к_файлу\имя_файла.xlsx</param>
        /// <returns></returns>
        public static bool checkWritingAvaliable(string xlFileName, Excel.Application xlApp)
        {
            bool result = false;
            Excel.Workbook workBook = xlApp.Workbooks.Open(xlFileName, false, true); // открываем наш файл на чтение
            try
            {
                using (var fs = File.Open(xlFileName, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    result = true;
                    //MessageBox.Show("файл свободен");
                }
            }
            catch 
            {
                result = false;
                //MessageBox.Show("файл занят");
            }
            workBook.Close(false, xlFileName);
            xlApp.Quit();
            return result;
        }

        /// <summary>
        /// Копирует диапазоны по top штук за раз
        /// </summary>
        /// <param name="ranges">Список диапазонов типа данных myExcelRange</param>
        /// <param name="top">По сколько диапазонов будет скопировано за раз</param>
        /// <param name="sheetCopy">Лист из которого будем копировать</param>
        /// <param name="sheetPaste">Лист в который будем вставлять</param>
        /// <param name="rowPasted">Строка в которую будем вставлять</param>
        /// <param name="columnPasted">Столбец в который будем вставлять</param>
        public static void CopyPasteRanges(List<myExcelRange> ranges, int top, Excel.Worksheet sheetCopy, Excel.Worksheet sheetPaste, int rowPasted, int columnPasted)
        {
            // строка, в которой через запятую перечислены диапазоны
            String rangeString = "";

            int start = 0;
            int end = ranges.Count - 1;
            int cc = 0;
            for (int i = 0; i < ranges.Count; i++, cc++)
            {
                if ((cc == top - 1) || i == ranges.Count - 1)
                {
                    end = i;
                    //richTextBox2.Text += "\n" + start + " - " + end + ":\n";
                    // тут пробежимся от старта до конца
                    for (int j = start; j <= end; j++)
                    {
                        if (j != end)
                        {
                            rangeString += ranges[j].adressString + ";";
                        }
                        else
                        {
                            rangeString += ranges[j].adressString;

                            //MessageBox.Show("Диапазон  (" + (end - start + 1) + "): " + rangeString + "\nВ строку: " + rowPasted);
                            // САМО КОПИРОВАНИЕ->ВСТАВКА
                            sheetCopy.Select();
                            sheetCopy.Range[rangeString].Copy();
                            sheetPaste.Select();
                            sheetPaste.Cells[rowPasted, columnPasted].PasteSpecial();

                            rangeString = "";
                            rowPasted += cc + 1;
                        }
                    }
                    //richTextBox2.Text += "\n";
                    start = end + 1;
                    cc = -1;
                }
            }
        }

        /// <summary>
        /// Копирует диапазоны построчно по top штук за раз
        /// </summary>
        /// <param name="ranges">Список диапазонов типа данных myExcelRange</param>
        /// <param name="top">По сколько строк будет скопировано за раз</param>
        /// <param name="sheetCopy">Лист из которого будем копировать</param>
        /// <param name="sheetPaste">Лист в который будем вставлять</param>
        /// <param name="rowPasted">Строка в которую будем вставлять</param>
        /// <param name="columnPasted">Столбец в который будем вставлять</param>
        public static void CopyPasteRows(List<myExcelRange> ranges, int top, Excel.Worksheet sheetCopy, Excel.Worksheet sheetPaste, int rowPasted, int columnPasted)
        {
            // строка, в которой через запятую перечислены диапазоны
            String rangeString = "";

            int start = 0;
            int end = ranges.Count - 1;
            int cc = 0;
            for (int i = 0; i < ranges.Count; i++, cc++)
            {
                if ((cc == top - 1) || i == ranges.Count - 1)
                {
                    end = i;
                    //richTextBox2.Text += "\n" + start + " - " + end + ":\n";
                    // тут пробежимся от старта до конца
                    for (int j = start; j <= end; j++)
                    {
                        if (j != end)
                        {
                            rangeString += ranges[j].adressRow + ";";
                        }
                        else
                        {
                            rangeString += ranges[j].adressRow;

                            //MessageBox.Show("Диапазон  (" + (end - start + 1) + "): " + rangeString + "\nВ строку: " + rowPasted);
                            // САМО КОПИРОВАНИЕ->ВСТАВКА
                            sheetCopy.Select();
                            sheetCopy.Range[rangeString].Copy();
                            sheetPaste.Select();
                            sheetPaste.Cells[rowPasted, columnPasted].PasteSpecial();

                            rangeString = "";
                            rowPasted += cc + 1;
                        }
                    }
                    //richTextBox2.Text += "\n";
                    start = end + 1;
                    cc = -1;
                }
            }
        }

        /// <summary>
        /// Возвращает список ячеек листа sheet из столбца columnLetter
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnLetter"></param>
        /// <returns></returns>
        public static List<String> getColumnStrings(Excel.Worksheet sheet, string columnLetter)
        {
            List<String> result = new List<String>();
            var idsAllArrayObjects = (object[,])sheet.Range[columnLetter + ":" + columnLetter].Value;
            foreach (var idObject in idsAllArrayObjects)
            {
                result.Add((String)idObject);
            }
            return result;
        }

        /// <summary>
        /// Возвращает двумерный массив значений диапазона myExcelRange range
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static object[,] getRange(Excel.Worksheet sheet, myExcelRange range)
        {
            return (object[,])sheet.Range[range.adressString].Value;
        }


        /// <summary>
        /// Получает на вход цвет в виде числа 16777215, возвращает нормальный RGB цвет
        /// </summary>
        /// <param name="colorInt"> Decimal представление цвета, напимер белый это 16777215</param>
        /// <returns></returns>
        public static Color getColorfromDecimal(int colorInt)
        {
            int B = colorInt % 256;
            int G = (colorInt - B) % (256 * 256);
            int R = (colorInt - B - G) % (256 * 256 * 256);

            R = R / (256 * 256);
            G = G / 256;
            B = B / 1;

            return Color.FromArgb(R, G, B); 
        }

        /// <summary>
        /// Заполняет dataGridView занчениями из листа workSheet
        /// </summary>
        /// <param name="datagridView">Заполняемый dataGridView</param>
        /// <param name="workSheet">Лист из которого будем заполнять</param>
        /// <param name="adrMaker">Собственный адресс мейкер</param>
        /// <param name="checkBox">ЧекБокс "Показать все" строки</param>
        public static void setDataGridView (DataGridView datagridView , Excel.Worksheet workSheet, myExcelAdressMaker adrMaker, System.Windows.Forms.CheckBox checkBox)
        {
            // Очистка памяти
            GC.Collect();


            // Количество строк и столбцов в новом листе
            int rowsCount = myExcel.getRowsCount(workSheet);
            int columnsCount = myExcel.getColumnsCount(workSheet);

            checkBox.Text = "Показать все " + rowsCount + " строк";

            // Получаем содержимое листа (за одно считытвание, фантастика)
            object[,] arrRange = (object[,])workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowsCount, columnsCount]].Value;//myExcel.getRange(workSheet, new myExcelRange(1, 1, rowsCount, columnsCount, adrMaker));

            // ДОБАВИМ СТОЛБЦЫ
            datagridView.Columns.Clear();
            // создадим массив со столбцами, а потом разом добавим их все на в dataGridView
            int tmpColumnsCount = columnsCount < 13 ? 14 : columnsCount;
            DataGridViewTextBoxColumn tmpColumn = new DataGridViewTextBoxColumn();
            DataGridViewColumn[] columns = new DataGridViewColumn[tmpColumnsCount];
            for (int i = 0; i < tmpColumnsCount; i++)
            {
                tmpColumn = new DataGridViewTextBoxColumn();
                tmpColumn.HeaderText = adrMaker.getLetter(i + 1);
                // зададим ширину столбца из реального экселя в dataGridView
                if (i < columnsCount)
                {
                    tmpColumn.Width = ((int)(workSheet.Columns[i + 1].ColumnWidth * 60)) / 10;
                    tmpColumn.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                columns[i] = tmpColumn;
            }
            datagridView.Columns.AddRange(columns);

            // ДОБАВИМ СТРОКИ И ЗАПОЛНИМ ЯЧЕЙКИ
            DataGridViewRow tmpRow = new DataGridViewRow();
            DataGridViewRow[] rows = new DataGridViewRow[rowsCount];
            if (!checkBox.Checked)
            {
                rowsCount = rowsCount > 11 ? 10 : rowsCount;
            }
            for (int i = 0; i < rowsCount; i++)
            {
                datagridView.Rows.Add();
                datagridView.Rows[i].HeaderCell.Value = (i + 1).ToString();

                for (int j = 0; j < columnsCount; j++)
                {
                    if (arrRange != null && arrRange[i + 1, j + 1] != null)
                    {
                        datagridView.Rows[i].Cells[j].Value = arrRange[i + 1, j + 1].ToString();
                    }
                }
            }
            // попытка очистить память
            arrRange = null;
        }

        public static void copySheetFromFileToFile (string fileFrom, string fileTo, string pastedSheetName)
        {
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook workBook1; //рабочая книга откуда будем копировать лист  
            Excel.Workbook workBook2; //рабочая книга куда будем копировать лист
            Excel.Worksheet sheet; //лист Excel            

            workBook1 = xlApp.Workbooks.Open(fileFrom); // название файла Excel откуда будем копировать лист
            workBook2 = xlApp.Workbooks.Open(fileTo); // название файла Excel куда будем копировать лист
            workBook1.Worksheets[1].Name = pastedSheetName;  // название листа который будем копировать
            sheet = workBook1.Worksheets[1];    // лист, который будем копировать
            sheet.Copy(After: workBook2.Worksheets[workBook2.Worksheets.Count]);  // сам процесс копирования листа из одного файла в другой           
            workBook2.Close(true); // закрываем и сохраняем изменения в файле 2    
            workBook1.Close(true); // закрываем и сохраняем изменения в файле 1   
            xlApp.Quit(); // закрываем Excel

            File.Delete(fileFrom);
        }





    }

}


// ЕСЛИ НУЖНО БУДЕТ КОГДА-НИБУДЬ ЗАПОЛНЯТЬ dataGridView ЗНАЧЕНИЯМИ ИЗ ТАБЛИЦЫ
///// <summary>
///// ОБвновление содержимого ДатаГридВью
///// </summary>
///// <param name="sender"></param>
///// <param name="e"></param>
//private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
//{
//    try
//    {
//        // Очистка памяти
//        GC.Collect();

//        richTextBox2.AppendText("Выбран " + listBox1.SelectedItem + "\n");
//        richTextBox2.AppendText("Формирование таблицы для предосмотра... " + "\n");
//        richTextBox2.ScrollToCaret();

//        // Текущий лист для предосмотра
//        Excel.Worksheet workSheet = workBook.Sheets[listBox1.SelectedIndex + 1];

//        // Количество строк и столбцов в новом листе
//        int rowsCount = myExcel.getRowsCount(workSheet);
//        int columnsCount = myExcel.getColumnsCount(workSheet);

//        // Получаем содержимое листа (за одно считытвание, фантастика)
//        object[,] arr = myExcel.getRange(workSheet, new myExcelRange(1, 1, rowsCount, columnsCount, adrMaker));


//        // ДОБАВИМ СТОЛБЦЫ
//        dataGridView1.Columns.Clear();
//        // создадим массив со столбцами, а потом разом добавим их все на в dataGridView
//        int tmpColumnsCount = columnsCount < 13 ? 14 : columnsCount;
//        DataGridViewTextBoxColumn tmpColumn = new DataGridViewTextBoxColumn();
//        DataGridViewColumn[] columns = new DataGridViewColumn[tmpColumnsCount];
//        for (int i = 0; i < tmpColumnsCount; i++)
//        {
//            tmpColumn = new DataGridViewTextBoxColumn();
//            tmpColumn.HeaderText = adrMaker.getLetter(i + 1);
//            // зададим ширину столбца из реального экселя в dataGridView
//            if (i < columnsCount)
//            {
//                tmpColumn.Width = ((int)(workSheet.Columns[i + 1].ColumnWidth * 60)) / 10;
//            }
//            columns[i] = tmpColumn;
//        }
//        this.dataGridView1.Columns.AddRange(columns);

//        // ДОБАВИМ СТРОКИ И ЗАПОЛНИМ ЯЧЕЙКИ
//        Excel.Range range = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[rowsCount, columnsCount]];
//        range.CopyPicture();

//        Image returnImage = Clipboard.GetImage();

//        DataGridViewRow tmpRow = new DataGridViewRow();
//        DataGridViewRow[] rows = new DataGridViewRow[rowsCount];
//        if (!checkBox3.Checked)
//        {
//            rowsCount = rowsCount > 18 ? 17 : rowsCount;
//        }
//        double tmpBackColor;
//        for (int i = 0; i < rowsCount; i++)
//        {
//            dataGridView1.Rows.Add();
//            dataGridView1.Rows[i].HeaderCell.Value = (i + 1).ToString();
//            for (int j = 0; j < columnsCount; j++)
//            {
//                if (arr != null && arr[i + 1, j + 1] != null)
//                {
//                    dataGridView1.Rows[i].Cells[j].Value = arr[i + 1, j + 1].ToString();
//                }
//            }
//        }


//        // закраска фонами
//        for (int i = 0; i < rowsCount; i++)
//        {
//            for (int j = 0; j < columnsCount; j++)
//            {
//                if (arr != null && arr[i + 1, j + 1] != null)
//                {
//                    tmpBackColor = range.Cells[i + 1, j + 1].Interior.Color;
//                    //richTextBox2.AppendText(tmpBackColor + "\n");
//                    if (tmpBackColor != 16777215)
//                    {
//                        dataGridView1.Rows[i].Cells[j].Style.BackColor = myExcel.getColorfromDecimal((int)tmpBackColor);
//                    }
//                }
//            }
//        }
//        // попытка очистить память
//        arr = null;
//    }
//    catch (Exception exc)
//    {
//        MessageBox.Show("Во время предосмотра листа произошла ошибка.\n" + exc.Message + "\n\nОбратитесь к Разработчику.\nРекомендуется не сохранять изменения в файле.");
//    }
//    finally
//    {
//        richTextBox2.AppendText("Таблица для предосмотра сформирована. " + "\n");
//        richTextBox2.ScrollToCaret();
//    }
//}

