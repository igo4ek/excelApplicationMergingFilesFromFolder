using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace excelApplicationFindAndCopy
{
    class myDelegates
    {
        // делегаторский класс -- обновлятор всея ричТекстБоксов
        public delegate void updateRichTextBox(String text, RichTextBox richTextBox);
        public updateRichTextBox richTextBoxUpdater;

        // делегаторский класс -- обновлятор активатор/дективатор всех кнопок
        public delegate void buttonCheckedChange(bool status, Button button);
        public buttonCheckedChange buttonCheckedChanger;

        // делегаторский класс -- обновлятор содержимого
        public delegate void listBoxUpdateContent(ListBox listBox, Excel.Workbook workBook);
        public listBoxUpdateContent listBoxUpdater;

        // делегаторский класс -- обновлятор активатор/дективатор всех кнопок
        public delegate void listBoxEnabledUpdate(ListBox listBox, bool status);
        public listBoxEnabledUpdate listBoxEnabledUpdater;

        // изменение listBox.selectedIndex
        public delegate void listBoxSelectedIndexChange(ListBox listBox, int selectedIndex);
        public listBoxSelectedIndexChange listBoxSelectedIndexChanger;


        public myDelegates()
        {
            // назначили делегаторскому объекту его метод по обновлению ричТекстБокса
            richTextBoxUpdater = (text, richTextBox) =>
            {
                richTextBox.AppendText(text);
                richTextBox.ScrollToCaret();
            };

            //назначили делегаторскому объекту по изменения статуса активности кнопки его метод
            buttonCheckedChanger = (status, button) =>
            {
                button.Enabled = status;
            };

            // (активация) включение или выключение листбокса
            listBoxEnabledUpdater = (listBox, status) =>
            {
                listBox.Enabled = status;
            };

            // Обновление содержимого listBox
            listBoxUpdater = (ListBox listBox, Excel.Workbook workBook) =>
            {
                try
                {
                    listBox.Items.Clear();
                    foreach (Excel.Worksheet sheet in workBook.Sheets)
                    {
                        listBox.Items.Add(sheet.Name + " (" + myExcel.getRowsCount(sheet) + ")");
                    }
                    listBox.SelectedIndex = 0;
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Во время открытия файла в DataGridView произошел сбой.\n" + exc.Message + "\n\nОбратитесь к Разработчику.\nРекомендуется не сохранять изменения в файле.");
                }
                finally
                {

                }
            };

            listBoxSelectedIndexChanger = (ListBox listBox, int selectedIndex) =>
            {
                listBox.SelectedIndex = selectedIndex;
            };
        }
    }
}
