using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ColumnName.Items.Add("Тарелка");
            ColumnName.Items.Add("Чашка");
            ColumnName.Items.Add("Ложка");
            ColumnName.Items.Add("Вилка");

            ColumnPrice.Items.Add("199,90");
            ColumnPrice.Items.Add("149,90");
            ColumnPrice.Items.Add("119,90");
            ColumnPrice.Items.Add("100,90");

            ColumnCode.Items.Add("501 111");
            ColumnCode.Items.Add("501 112");
            ColumnCode.Items.Add("601 211");
            ColumnCode.Items.Add("601 212");
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker2.Value == dateTimePicker3.Value)
            {
                MessageBox.Show("Одинаковые даты в отчетном периоде", "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
            }
            int check = DateTime.Compare(dateTimePicker3.Value.Date, dateTimePicker2.Value.Date);
            if (check < 0)
            {
                MessageBox.Show("Даты не подходят по периоду", "Предупреждение",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

            }
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker3.Value == dateTimePicker2.Value)
            {
                MessageBox.Show("Одинаковые даты в отчетном периоде", "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
            }
            int check = DateTime.Compare(dateTimePicker3.Value.Date, dateTimePicker2.Value.Date);
            if (check < 0)
            {
                MessageBox.Show("Даты не подходят по периоду", "Предупреждение",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

            }
        }

        private void comboBox1_ValueMemberChanged(object sender, EventArgs e)
        {
            if (comboBox_Organisation.ValueMember == "ООО Вкусная жизнь")
            {
                MessageBox.Show("Одинаковые даты в отчетном периоде", "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
            }
        }
        private void comboBox_remove(ComboBox cm1, ComboBox cm2, ComboBox cm3, string rem)
        {
            cm1.Items.Remove(rem);
            cm2.Items.Remove(rem);
            cm3.Items.Remove(rem);
        }
        private void comboBox_add(ComboBox cm1, ComboBox cm2, ComboBox cm3, string add)
        {
            cm1.Items.Add(add);
            cm2.Items.Add(add);
            cm3.Items.Add(add);
        }
        private void comboBox_text(ComboBox cm1, ComboBox cm2, ComboBox cm3, string text)
        {
            cm1.Text = text;
            cm2.Text = text;
            cm3.Text = text;
        }
        private void comboboxRole(ComboBox cm1, ComboBox decode1)
        {
            if (comboBox_Place.Text == "Столовая Вкуснятина")
            {
                if (cm1.Text == "Главный секретарь")
                {
                    decode1.Text = "Афанасьева Вероника Владимировна";
                }
                if (cm1.Text == "Секретарь")
                {
                    decode1.Text = "Макарова Аврора Никитична";
                }
                if (cm1.Text == "Заведующий складом")
                {
                    decode1.Text = "Иванова Юлия Ивановна";
                }
            }
            if (comboBox_Place.Text == "Кафе Сказка")
            {
                if (cm1.Text == "Главный секретарь")
                {
                    decode1.Text = "Нечаева Ника Михайловна";
                }
                if (cm1.Text == "Секретарь")
                {
                    decode1.Text = "Яковлев Михаил Фёдорович";
                }
                if (cm1.Text == "Заведующий складом")
                {
                    decode1.Text = "Петрова Александра Кирилловна";
                }
            }
            if (comboBox_Place.Text == "Ресторан Престиж")
            {
                if (cm1.Text == "Главный секретарь")
                {
                    decode1.Text = "Новиков Георгий Маркович";
                }
                if (cm1.Text == "Секретарь")
                {
                    decode1.Text = "Виноградова Василиса Никитична";
                }
                if (cm1.Text == "Заведующий складом")
                {
                    decode1.Text = "Фролов Иван Матвеевич";
                }
            }
            if (comboBox_Decode1.Text != "" && comboBox_Decode2.Text != "" && comboBox_Decode3.Text!= "")
            {
                if (comboBox_Decode1.Text == comboBox_Decode2.Text || comboBox_Decode1.Text == comboBox_Decode3.Text || comboBox_Decode2.Text == comboBox_Decode3.Text)
                {
                    MessageBox.Show("Повторяющиеся члены комиссии", "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                    cm1.Text = "";
                    decode1.Text = "";
                }
            }
        }

        private void comboBoxDecode(ComboBox cm1, ComboBox decode1)
        {
            if (comboBox_Place.Text == "Столовая Вкуснятина")
            {
                if (decode1.Text == "Афанасьева Вероника Владимировна")
                {
                    cm1.Text = "Главный секретарь";
                }
                if (decode1.Text == "Макарова Аврора Никитична")
                {
                    cm1.Text = "Секретарь";
                }
                if (decode1.Text == "Иванова Юлия Ивановна")
                {
                    cm1.Text = "Заведующий складом";
                }
            }
            if (comboBox_Place.Text == "Кафе Сказка")
            {
                if (decode1.Text == "Нечаева Ника Михайловна")
                {
                    cm1.Text = "Главный секретарь";
                }
                if (decode1.Text == "Яковлев Михаил Фёдорович")
                {
                    cm1.Text = "Секретарь";
                }
                if (decode1.Text == "Петрова Александра Кирилловна")
                {
                    cm1.Text = "Заведующий складом";
                }
            }
            if (comboBox_Place.Text == "Ресторан Престиж")
            {
                if (decode1.Text == "Новиков Георгий Маркович")
                {
                    cm1.Text = "Главный секретарь";
                }
                if (decode1.Text == "Виноградова Василиса Никитична")
                {
                    cm1.Text = "Секретарь";
                }
                if (decode1.Text == "Фролов Иван Матвеевич")
                {
                    cm1.Text = "Заведующий складом";
                }
            }
            if (comboBox_Decode1.Text != "" && comboBox_Decode2.Text != "" && comboBox_Decode3.Text != "")
            {
                if (comboBox_Decode1.Text == comboBox_Decode2.Text || comboBox_Decode1.Text == comboBox_Decode3.Text || comboBox_Decode2.Text == comboBox_Decode3.Text)
                {
                    MessageBox.Show("Повторяющиеся члены комиссии", "Предупреждение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                    cm1.Text = "";
                    decode1.Text = "";
                }
            }
        }
        private void comboBox_Organisation_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Задается подразделение
            comboBox_Place.Items.Remove("Столовая Вкуснятина");
            comboBox_Place.Items.Remove("Кафе Сказка");
            comboBox_Place.Items.Remove("Ресторан Престиж");
            if (comboBox_Organisation.Text == "ООО Вкусная жизнь")
            {
                comboBox_Place.Items.Add("Столовая Вкуснятина");
                comboBox_Place.Text = "Столовая Вкуснятина";
            }
            if (comboBox_Organisation.Text == "ООО Сказка вкуса")
            {
                comboBox_Place.Items.Add("Кафе Сказка");
                comboBox_Place.Text = "Кафе Сказка";
            }
            if (comboBox_Organisation.Text == "ООО Высокий стандарт")
            {
                comboBox_Place.Items.Add("Ресторан Престиж");
                comboBox_Place.Text = "Ресторан Престиж";
            }
            //Задается комиссия
            comboBox_remove(comboBox_Role1, comboBox_Role2, comboBox_Role3, "Секретарь");
            comboBox_remove(comboBox_Role1, comboBox_Role2, comboBox_Role3, "Главный секретарь");
            comboBox_remove(comboBox_Role1, comboBox_Role2, comboBox_Role3, "Заведующий складом");

            comboBox_remove(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Афанасьева Вероника Владимировна");
            comboBox_remove(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Макарова Аврора Никитична");
            comboBox_remove(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Иванова Юлия Ивановна");

            comboBox_remove(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Нечаева Ника Михайловна");
            comboBox_remove(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Яковлев Михаил Фёдорович");
            comboBox_remove(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Петрова Александра Кирилловна");

            comboBox_remove(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Новиков Георгий Маркович");
            comboBox_remove(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Виноградова Василиса Никитична");
            comboBox_remove(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Фролов Иван Матвеевич");

            comboBox_text(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "");
            comboBox_text(comboBox_Role1, comboBox_Role2, comboBox_Role3, "");

            comboBox_add(comboBox_Role1, comboBox_Role2, comboBox_Role3, "Секретарь");
            comboBox_add(comboBox_Role1, comboBox_Role2, comboBox_Role3, "Главный секретарь");
            comboBox_add(comboBox_Role1, comboBox_Role2, comboBox_Role3, "Заведующий складом");

            if (comboBox_Place.Text == "Столовая Вкуснятина")
            {
                comboBox_add(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Афанасьева Вероника Владимировна");
                comboBox_add(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Макарова Аврора Никитична");
                comboBox_add(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Иванова Юлия Ивановна");
            }
            if (comboBox_Place.Text == "Кафе Сказка")
            {
                comboBox_add(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Нечаева Ника Михайловна");
                comboBox_add(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Яковлев Михаил Фёдорович");
                comboBox_add(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Петрова Александра Кирилловна");
            }
            if (comboBox_Place.Text == "Ресторан Престиж")
            {
                comboBox_add(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Новиков Георгий Маркович");
                comboBox_add(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Виноградова Василиса Никитична");
                comboBox_add(comboBox_Decode1, comboBox_Decode2, comboBox_Decode3, "Фролов Иван Матвеевич");
            }
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            int check = DateTime.Compare(dateTimePicker4.Value.Date, dateTimePicker1.Value.Date);
            if(check < 0)
            {
                MessageBox.Show("Дата составления больше даты утверждения", "Предупреждение",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

            }
        }

        private void textBox_Material_Role_TextChanged(object sender, EventArgs e)
        {

        }

        private void button_Save_Click(object sender, EventArgs e)
        {
            //Создаем объекты для работы с excel
            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlwb;
            Microsoft.Office.Interop.Excel.Worksheet xlws;
            Microsoft.Office.Interop.Excel.Range xlr;

            object ms = System.Reflection.Missing.Value;
            xls.Visible = true;
            xlwb = xls.Workbooks.Add("OP-8.xls");
            xlws = xlwb.Worksheets.get_Item(1);
            xlr = xlws.Cells[1, 1];
            //Выводим данные в excel

            xlws.Cells[14, 37] = textBox1.Text;
            xlws.Cells[14, 45] = dateTimePicker1.Value.ToShortDateString();
            xlws.Cells[14, 53] = dateTimePicker2.Value.Date.ToShortDateString();
            xlws.Cells[14, 58] = dateTimePicker3.Value.Date.ToShortDateString();
            xlws.Cells[14, 58] = dateTimePicker3.Value.Date.ToShortDateString();
            xlws.Cells[6, 1] = comboBox_Organisation.Text;
            xlws.Cells[8, 1] = comboBox_Place.Text;

            xlws.Cells[13, 65] = comboBox_Approved.Text;

            xlws.Cells[17, 64] = dateTimePicker4.Value.Day.ToString();
            xlws.Cells[17, 66] = dateTimePicker4.Value.Month.ToString();
            xlws.Cells[17, 74] = dateTimePicker4.Value.Year.ToString();

            xlws.Cells[18, 21] = textBox_Material_Role.Text;
            xlws.Cells[18, 39] = comboBox_Material.Text;

            xlws.Cells[65, 14] = comboBox_Role1.Text;
            xlws.Cells[67, 14] = comboBox_Role2.Text;
            xlws.Cells[69, 14] = comboBox_Role3.Text;

            xlws.Cells[65, 43] = comboBox_Decode1.Text;
            xlws.Cells[67, 43] = comboBox_Decode2.Text;
            xlws.Cells[69, 43] = comboBox_Decode3.Text;

            xlws.Cells[38, 26] = textBoxItogBreak.Text;

            double resBroken = 0, resLost = 0;

            xlws.Cells[38, 35] = textBoxItogLost.Text;

            xlws.Cells[38, 44] = textBoxItogTotal.Text;
            xlws.Cells[38, 48] = textBoxItogCost.Text;

            for (int j = 0; j < dataGridView1.RowCount - 1; j++)
            {
                xlws.Cells[28 + j, 1] = dataGridView1.Rows[j].Cells[0].Value;
                xlws.Cells[28 + j, 5] = dataGridView1.Rows[j].Cells[1].Value;
                xlws.Cells[28 + j, 15] = dataGridView1.Rows[j].Cells[2].Value;
                xlws.Cells[28 + j, 20] = dataGridView1.Rows[j].Cells[3].Value;
                xlws.Cells[28 + j, 26] = dataGridView1.Rows[j].Cells[4].Value;
                xlws.Cells[28 + j, 30] = Convert.ToDouble(dataGridView1.Rows[j].Cells[3].Value) * Convert.ToDouble(dataGridView1.Rows[j].Cells[4].Value);
                xlws.Cells[28 + j, 35] = dataGridView1.Rows[j].Cells[5].Value;
                xlws.Cells[28 + j, 39] = Convert.ToDouble(dataGridView1.Rows[j].Cells[3].Value) * Convert.ToDouble(dataGridView1.Rows[j].Cells[5].Value);
                xlws.Cells[28 + j, 44] = dataGridView1.Rows[j].Cells[6].Value;
                xlws.Cells[28 + j, 48] = dataGridView1.Rows[j].Cells[7].Value;
                xlws.Cells[28 + j, 53] = dataGridView1.Rows[j].Cells[8].Value;
                xlws.Cells[28 + j, 70] = dataGridView1.Rows[j].Cells[9].Value;
                resBroken += Convert.ToDouble(dataGridView1.Rows[j].Cells[3].Value) * Convert.ToDouble(dataGridView1.Rows[j].Cells[4].Value);
                resLost += Convert.ToDouble(dataGridView1.Rows[j].Cells[3].Value) * Convert.ToDouble(dataGridView1.Rows[j].Cells[5].Value);
            }

            xlws.Cells[38, 30] = Convert.ToString(resBroken);
            xlws.Cells[38, 39] = Convert.ToString(resLost);
            xlwb.Save();
            xlwb.Close();
            xls.Quit();
        }

        private void button_Export_Click(object sender, EventArgs e)
        {
            //Создаем объекты для работы с excel
            bool check = copyToClipboard();
            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlwb;
            Microsoft.Office.Interop.Excel.Worksheet xlws;
            Microsoft.Office.Interop.Excel.Range xlr;

            object ms = System.Reflection.Missing.Value;
            xls.Visible = true;
            xlwb = xls.Workbooks.Add("OP-8.xls");
            xlws = xlwb.Worksheets.get_Item(1);
            xlr = xlws.Cells[1, 1];
            //Выводим данные в excel

            xlws.Cells[14,37] = textBox1.Text;
            xlws.Cells[14, 45] = dateTimePicker1.Value.ToShortDateString();
            xlws.Cells[14, 53] = dateTimePicker2.Value.Date.ToShortDateString();
            xlws.Cells[14, 58] = dateTimePicker3.Value.Date.ToShortDateString();
            xlws.Cells[14, 58] = dateTimePicker3.Value.Date.ToShortDateString();
            xlws.Cells[6, 1] = comboBox_Organisation.Text;
            xlws.Cells[8, 1] = comboBox_Place.Text;

            xlws.Cells[13, 65] = comboBox_Approved.Text;

            xlws.Cells[17, 64] = dateTimePicker4.Value.Day.ToString();
            xlws.Cells[17, 66] = dateTimePicker4.Value.Month.ToString();
            xlws.Cells[17, 74] = dateTimePicker4.Value.Year.ToString();

            xlws.Cells[18, 21] = textBox_Material_Role.Text;
            xlws.Cells[18, 39] = comboBox_Material.Text;

            xlws.Cells[65, 14] = comboBox_Role1.Text;
            xlws.Cells[67, 14] = comboBox_Role2.Text;
            xlws.Cells[69, 14] = comboBox_Role3.Text;

            xlws.Cells[65, 43] = comboBox_Decode1.Text;
            xlws.Cells[67, 43] = comboBox_Decode2.Text;
            xlws.Cells[69, 43] = comboBox_Decode3.Text;

            xlws.Cells[38, 26] = textBoxItogBreak.Text;

            double resBroken=0, resLost=0;

            xlws.Cells[38, 35] = textBoxItogLost.Text;

            xlws.Cells[38, 44] = textBoxItogTotal.Text;
            xlws.Cells[38, 48] = textBoxItogCost.Text;


            if (check)
            {
                for (int j = 0; j < dataGridView1.RowCount - 1; j++)
                {
                    xlws.Cells[28 + j, 1] = dataGridView1.Rows[j].Cells[0].Value;
                    xlws.Cells[28 + j, 5] = dataGridView1.Rows[j].Cells[1].Value;
                    xlws.Cells[28 + j, 15] = dataGridView1.Rows[j].Cells[2].Value;
                    xlws.Cells[28 + j, 20] = dataGridView1.Rows[j].Cells[3].Value;
                    xlws.Cells[28 + j, 26] = dataGridView1.Rows[j].Cells[4].Value;
                    xlws.Cells[28 + j, 30] = Convert.ToDouble(dataGridView1.Rows[j].Cells[3].Value) * Convert.ToDouble(dataGridView1.Rows[j].Cells[4].Value);
                    xlws.Cells[28 + j, 35] = dataGridView1.Rows[j].Cells[5].Value;
                    xlws.Cells[28 + j, 39] = Convert.ToDouble(dataGridView1.Rows[j].Cells[3].Value) * Convert.ToDouble(dataGridView1.Rows[j].Cells[5].Value);
                    xlws.Cells[28 + j, 44] = dataGridView1.Rows[j].Cells[6].Value;
                    xlws.Cells[28 + j, 48] = dataGridView1.Rows[j].Cells[7].Value;
                    xlws.Cells[28 + j, 53] = dataGridView1.Rows[j].Cells[8].Value;
                    xlws.Cells[28 + j, 70] = dataGridView1.Rows[j].Cells[9].Value;
                    resBroken += Convert.ToDouble(dataGridView1.Rows[j].Cells[3].Value) * Convert.ToDouble(dataGridView1.Rows[j].Cells[4].Value);
                    resLost += Convert.ToDouble(dataGridView1.Rows[j].Cells[3].Value) * Convert.ToDouble(dataGridView1.Rows[j].Cells[5].Value);
                }

            }
            xlws.Cells[38, 30] = Convert.ToString(resBroken);
            xlws.Cells[38, 39] = Convert.ToString(resLost);
            xlwb.Save();
            xlwb.Close();
            xls.Quit();
        }

        private bool copyToClipboard()
        {
            dataGridView1.SelectAll();
            DataObject data = dataGridView1.GetClipboardContent();
            if (data != null)
            {
                Clipboard.SetDataObject(data);
                return true;
            }
            else
                MessageBox.Show("Пустая таблица");
            return false;
        }
        private void button_NewDoc_Click(object sender, EventArgs e)
        {
            //Очистка данных
            textBox1.Text = "";
            dateTimePicker1.Value = DateTime.Today.AddDays(-3);
            dateTimePicker2.Value = DateTime.Today.AddDays(-2);
            dateTimePicker3.Value = DateTime.Today.AddDays(-1);
            dateTimePicker4.Value = DateTime.Today.AddDays(0);
            comboBox_Organisation.Text = "";
            comboBox_Place.Text = "";
            comboBox_Material.Text = "";
            textBox_Material_Role.Text = "";
            comboBox_Decode1.Text = "";
            comboBox_Decode2.Text = "";
            comboBox_Decode3.Text = "";
            comboBox_Role1.Text = "";
            comboBox_Role2.Text = "";
            comboBox_Role3.Text = "";
            textBoxItogBreak.Text = "";
            textBoxItogLost.Text = "";
            textBoxItogTotal.Text = "";
            textBoxItogCost.Text = "";
            while(dataGridView1.Rows.Count > 1)
                dataGridView1.Rows.RemoveAt(0);
        }

        private void comboBox_Place_TextChanged(object sender, EventArgs e)
        {

            //Задается утверждающее лицо
            comboBox_Approved.Items.Remove("Иванова Юлия Ивановна");
            comboBox_Approved.Items.Remove("Макарова Аврора Никитична");
            comboBox_Approved.Items.Remove("Нечаева Ника Михайловна");
            comboBox_Approved.Items.Remove("Голубев Арсений Дмитриевич");
            comboBox_Approved.Items.Remove("Петрова Александра Кирилловна");
            comboBox_Approved.Items.Remove("Лобанова Ева Дмитриевна");
            comboBox_Approved.Items.Remove("Новиков Георгий Маркович");
            comboBox_Approved.Items.Remove("Акимов Андрей Фёдорович");

            comboBox_Approved.Text = "";

            if (comboBox_Place.Text == "Столовая Вкуснятина")
            {
                comboBox_Approved.Items.Add("Иванова Юлия Ивановна");
                comboBox_Approved.Items.Add("Макарова Аврора Никитична");
            }
            if (comboBox_Place.Text == "Кафе Сказка")
            {
                comboBox_Approved.Items.Add("Нечаева Ника Михайловна");
                comboBox_Approved.Items.Add("Голубев Арсений Дмитриевич");
                comboBox_Approved.Items.Add("Петрова Александра Кирилловна");
            }
            if (comboBox_Place.Text == "Ресторан Престиж")
            {
                comboBox_Approved.Items.Add("Лобанова Ева Дмитриевна");
                comboBox_Approved.Items.Add("Новиков Георгий Маркович");
                comboBox_Approved.Items.Add("Акимов Андрей Фёдорович");
            }

            //Задается материально ответственное лицо
            comboBox_Material.Items.Remove("Афанасьева Вероника Владимировна");
            comboBox_Material.Items.Remove("Макарова Аврора Никитична");
            comboBox_Material.Items.Remove("Иванова Юлия Ивановна");
            comboBox_Material.Items.Remove("Нечаева Ника Михайловна");
            comboBox_Material.Items.Remove("Яковлев Михаил Фёдорович");
            comboBox_Material.Items.Remove("Петрова Александра Кирилловна");
            comboBox_Material.Items.Remove("Новиков Георгий Маркович");
            comboBox_Material.Items.Remove("Виноградова Василиса Никитична");
            comboBox_Material.Items.Remove("Фролов Иван Матвеевич");

            comboBox_Material.Text = "";
            if (comboBox_Place.Text == "Столовая Вкуснятина")
            {
                comboBox_Material.Items.Add("Афанасьева Вероника Владимировна");
                comboBox_Material.Items.Add("Макарова Аврора Никитична");
                comboBox_Material.Items.Add("Иванова Юлия Ивановна");
            }
            if (comboBox_Place.Text == "Кафе Сказка")
            {
                comboBox_Material.Items.Add("Нечаева Ника Михайловна");
                comboBox_Material.Items.Add("Яковлев Михаил Фёдорович");
                comboBox_Material.Items.Add("Петрова Александра Кирилловна");
            }
            if (comboBox_Place.Text == "Ресторан Престиж")
            {
                comboBox_Material.Items.Add("Новиков Георгий Маркович");
                comboBox_Material.Items.Add("Виноградова Василиса Никитична");
                comboBox_Material.Items.Add("Фролов Иван Матвеевич");
            }
        }

        private void comboBoxMaterial_SelectedIndexChanged(object sender, EventArgs e)  //Определение роли по ФИО
        {
            textBox_Material_Role.Text = "";
            if (comboBox_Material.Text == "Афанасьева Вероника Владимировна" || comboBox_Material.Text == "Нечаева Ника Михайловна" ||
                comboBox_Material.Text == "Новиков Георгий Маркович")
            {
                textBox_Material_Role.Text = "Главный секретать";
            }
            if (comboBox_Material.Text == "Макарова Аврора Никитична" || comboBox_Material.Text == "Яковлев Михаил Фёдорович" ||
    comboBox_Material.Text == "Виноградова Василиса Никитична")
            {
                textBox_Material_Role.Text = "Секретарь";
            }
            if (comboBox_Material.Text == "Иванова Юлия Ивановна" || comboBox_Material.Text == "Петрова Александра Кирилловна" ||
    comboBox_Material.Text == "Фролов Иван Матвеевич")
            {
                textBox_Material_Role.Text = "Заведующий складом";
            }
        }

        private void comboBoxDecode1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxDecode(comboBox_Role1, comboBox_Decode1);
        }

        private void comboBox_Decode2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxDecode(comboBox_Role2, comboBox_Decode2);
        }

        private void comboBox_Decode3_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBoxDecode(comboBox_Role3, comboBox_Decode3);
        }

        private void comboBox_Role2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboboxRole(comboBox_Role2, comboBox_Decode2);
        }

        private void comboBox_Role3_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboboxRole(comboBox_Role3, comboBox_Decode3);
        }

        private void comboBox_Role1_SelectedIndexChanged(object sender, EventArgs e)    //Определение ФИО от роли для первой подписи комиссии
        {
            comboboxRole(comboBox_Role1, comboBox_Decode1);
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            int check = DateTime.Compare(dateTimePicker4.Value.Date, dateTimePicker1.Value.Date);
            if (check < 0)
            {
                MessageBox.Show("Дата составления больше даты утверждения", "Предупреждение",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

            }
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e) //Взаимосвязь ячеек в таблице
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == 1)
            {
                var selectedValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (selectedValue != null)
                {
                    switch (selectedValue.ToString())
                    {
                        case "Тарелка":
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = "501 111";
                            dataGridView1.Rows[e.RowIndex].Cells[3].Value = "199,90";
                            break;
                        case "Чашка":
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = "501 112";
                            dataGridView1.Rows[e.RowIndex].Cells[3].Value = "149,90";
                            break;
                        case "Вилка":
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = "601 212";
                            dataGridView1.Rows[e.RowIndex].Cells[3].Value = "100,90";
                            break;
                        case "Ложка":
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = "601 211";
                            dataGridView1.Rows[e.RowIndex].Cells[3].Value = "119,90";
                            break;
                    }
                }
            }
            if (e.RowIndex >= 0 && e.ColumnIndex == 2)
            {
                var selectedValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (selectedValue != null)
                {
                    switch (selectedValue.ToString())
                    {
                        case "501 111":
                            dataGridView1.Rows[e.RowIndex].Cells[1].Value = "Тарелка";
                            dataGridView1.Rows[e.RowIndex].Cells[3].Value = "199,90";
                            break;
                        case "501 112":
                            dataGridView1.Rows[e.RowIndex].Cells[1].Value = "Чашка";
                            dataGridView1.Rows[e.RowIndex].Cells[3].Value = "149,90";
                            break;
                        case "601 212":
                            dataGridView1.Rows[e.RowIndex].Cells[1].Value = "Вилка";
                            dataGridView1.Rows[e.RowIndex].Cells[3].Value = "100,90";
                            break;
                        case "601 211":
                            dataGridView1.Rows[e.RowIndex].Cells[1].Value = "Ложка";
                            dataGridView1.Rows[e.RowIndex].Cells[3].Value = "119,90";
                            break;
                    }
                }
            }
            if (e.RowIndex >= 0 && e.ColumnIndex == 3)
            {
                var selectedValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (selectedValue != null)
                {
                    switch (selectedValue.ToString())
                    {
                        case "199,90":
                            dataGridView1.Rows[e.RowIndex].Cells[1].Value = "Тарелка";
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = "501 111";
                            break;
                        case "149,90":
                            dataGridView1.Rows[e.RowIndex].Cells[1].Value = "Чашка";
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = "501 112";
                            break;
                        case "100,90":
                            dataGridView1.Rows[e.RowIndex].Cells[1].Value = "Вилка";
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = "601 212";
                            break;
                        case "119,90":
                            dataGridView1.Rows[e.RowIndex].Cells[1].Value = "Ложка";
                            dataGridView1.Rows[e.RowIndex].Cells[2].Value = "601 211";
                            break;
                    }
                }
            }
            if (e.RowIndex >= 0)
            {
                var selectedValue3 = dataGridView1.Rows[e.RowIndex].Cells[3].Value;
                var selectedValue4 = dataGridView1.Rows[e.RowIndex].Cells[4].Value;
                var selectedValue5 = dataGridView1.Rows[e.RowIndex].Cells[5].Value;
                if (selectedValue4 != null && selectedValue5 != null)
                {
                    int broken = Convert.ToInt32(selectedValue4);
                    int lost = Convert.ToInt32(selectedValue5);
                    string res = Convert.ToString(broken + lost);
                    dataGridView1.Rows[e.RowIndex].Cells[6].Value = res;
                    if (selectedValue3 != null)
                    {
                        double price = Convert.ToDouble(selectedValue3);
                        string TotalCost = Convert.ToString(price * (broken + lost));
                        dataGridView1.Rows[e.RowIndex].Cells[7].Value = TotalCost;

                        int itogBreak = 0, itogLost = 0;
                        double itogCost = 0;
                        for (int i = 0; i <= e.RowIndex; i++)
                        {
                            itogBreak += Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value);
                            itogLost += Convert.ToInt32(dataGridView1.Rows[i].Cells[5].Value);
                            itogCost += Convert.ToDouble(dataGridView1.Rows[i].Cells[7].Value);
                        }
                        textBoxItogBreak.Text = Convert.ToString(itogBreak);
                        textBoxItogLost.Text = Convert.ToString(itogLost);
                        textBoxItogTotal.Text = Convert.ToString(itogBreak + itogLost);
                        textBoxItogCost.Text = Convert.ToString(itogCost);
                    }
                }

            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged_1(object sender, EventArgs e)
        {
            dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
