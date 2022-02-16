using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace Bookstore.Framework
{
    public static class Utils
    {
        public static bool isTextBoxNull(TextBox box)
        {
            return (box.Text.Length > 0) ? true : false;
        }

        public static bool isMaskTextBoxNull(MaskedTextBox box)
        {
            return (box.Text.Replace(" ", "") == ",,") ? false : true;
        }

        public static bool isComboBoxNull(ComboBox box)
        {
            return (box.SelectedIndex > -1) ? true : false;
        }

        public static void addCombo(DataGridView dataGridView, ComboBox comboBox, int j)
        {
            comboBox.Items.Clear();
            for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
                comboBox.Items.Add(dataGridView[j, i].Value);
        }

        public static void addComboAuthor(DataGridView dataGridView, ComboBox comboBox)
        {
            comboBox.Items.Clear();
            for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
                comboBox.Items.Add(dataGridView[1, i].Value.ToString().Replace(" ", "") + " " + dataGridView[2, i].Value.ToString().Replace(" ", "") + " " + dataGridView[3, i].Value.ToString().Replace(" ", ""));
        }

        public static void addComboBuyer(DataGridView dataGridView, ComboBox comboBox)
        {
            comboBox.Items.Clear();
            for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
                comboBox.Items.Add(dataGridView[1, i].Value.ToString().Replace(" ", "") + " " + dataGridView[2, i].Value.ToString().Replace(" ", "") + " " + dataGridView[3, i].Value.ToString().Replace(" ", ""));
        }

        public static void addComboEmployee(DataGridView dataGridView, ComboBox comboBox)
        {
            comboBox.Items.Clear();
            for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
                comboBox.Items.Add(dataGridView[1, i].Value.ToString().Replace(" ", "") + " " + dataGridView[2, i].Value.ToString().Replace(" ", "") + " " + dataGridView[3, i].Value.ToString().Replace(" ", ""));
        }

        public static void FindRecord(DataGridView grid, TextBox textBox)
        {
            int count = 0;
            for (int i = 0; i < grid.Columns.Count - 1; i++)
            {
                for (int j = 0; j < grid.Rows.Count - 1; j++)
                {
                    if ((Convert.ToString(grid[i, j].Value).ToUpper()).Contains(textBox.Text.ToUpper()))
                    {
                        ++count;
                        grid.Rows[j].DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 70, 250);
                        grid.Rows[j].Cells[i].Style.BackColor = Color.FromArgb(255, 250, 70, 180);
                    }
                }
            }
            switch (count)
            {
                case 1:
                    MessageBox.Show($"Найдена {count} совпадений");
                    break;
                case 2:
                case 3:
                case 4:
                    MessageBox.Show($"Найдены {count} совпадений");
                    break;
                default:
                    MessageBox.Show($"Найдено {count} совпадений");
                    break;
            }
        }

        public static void ComboGrid(DataGridView grid, ComboBox comboBox)
        {
            comboBox.Items.Clear();
            for (int i = 1; i < grid.Columns.Count; i++)
            {
                comboBox.Items.Add(grid.Columns[i].HeaderText);
            }
        }

        public static void FilterGrid(DataGridView grid, ComboBox comboBox, TextBox textBox)
        {
            grid.CurrentCell = null;
            for (int i = 0; i < grid.Rows.Count - 1; i++)
                grid.Rows[i].Visible = false;
            for (int i = 0; i < grid.Columns.Count; i++)
                if (grid.Columns[i].HeaderText == Convert.ToString(comboBox.Items[comboBox.SelectedIndex]))
                    for (int j = 0; j < grid.Rows.Count - 1; j++)
                    {
                        if (textBox.Text == Convert.ToString(grid[i, j].Value).Replace(" ", ""))
                            grid.Rows[j].Visible = true;
                    }    
                        
        }

        public static void SaveTable(DataGridView grid, string name)
        {
            string path = System.IO.Directory.GetCurrentDirectory() + @"\" + $"{name}.xlsx";
            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook workbook = excelapp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            for (int i = 2; i < grid.RowCount + 2; i++)
                for (int j = 2; j < grid.ColumnCount+1; j++)
                {
                    worksheet.Rows[1].Columns[j-1] = grid.Columns[j - 1].HeaderText;
                    worksheet.Rows[i].Columns[j-1] = grid.Rows[i - 2].Cells[j - 1].Value;
                }


            excelapp.AlertBeforeOverwriting = false;
            workbook.SaveAs(path);
            excelapp.Quit();
            MessageBox.Show($"Таблица {name} сохранена в Excel");
        }

        public static void ChangeTab(TabControl tabControl, int i)
        {
            tabControl.SelectTab(i);
        }

        public static void WordBook(DataGridView grid1, DataGridView grid2)
        {
            MessageBox.Show("Данные из таблицы успешно сохранены в Word");
            var helper = new WordHelper("КнижныйПаспорт.docx");
            var Count = "";
            for (int i = 0; i < grid1.RowCount - 1; i++)
            {
                for (int j = 0; j < grid2.RowCount - 1; j++)
                {
                    if (Convert.ToString(grid1[5, i].Value) == Convert.ToString(grid2[1, j].Value))
                    {
                        Count = Convert.ToString(grid2[2, j].Value);
                    }
                }
                var items = new Dictionary<string, string>
                {
                    { "<Name>", Convert.ToString(grid1[5, i].Value) },
                    { "<Count>", Count },
                    { "<Author>", Convert.ToString(grid1[2, i].Value)},
                    { "<Genre>", Convert.ToString(grid1[1, i].Value)},
                    { "<Izd>", Convert.ToString(grid1[3, i].Value) },
                    { "<Lan>", Convert.ToString(grid1[4, i].Value) },
                    { "<Date>", Convert.ToString(grid1[6, i].Value) },
                    { "<Sale>", Convert.ToString(grid1[7, i].Value) },
                };
                helper.Process(items);
            }
        }

        public static void WordOrder(DataGridView grid1)
        {
            MessageBox.Show("Данные из таблицы успешно сохранены в Word");
            var helper = new WordHelper("Заказы.docx");
            for (int i = 0; i < grid1.RowCount - 1; i++)
            { 
                var items = new Dictionary<string, string>
                {
                    { "<Name>", Convert.ToString(grid1[1, i].Value) },
                    { "<Buyer>", Convert.ToString(grid1[2, i].Value)},
                    { "<Employee>", Convert.ToString(grid1[3, i].Value)},
                    { "<Discount>", Convert.ToString(grid1[4, i].Value)},
                    { "<Discount2>", Convert.ToString(grid1[5, i].Value) },
                    { "<Count>", Convert.ToString(grid1[6, i].Value) },
                    { "<Sale>", Convert.ToString(grid1[7, i].Value) },
                    { "<EndSale>", Convert.ToString(grid1[8, i].Value) },
                };
                helper.Process(items);
            }
        }

        public static void WordSupply(DataGridView grid1)
        {
            MessageBox.Show("Данные из таблицы успешно сохранены в Word");
            var helper = new WordHelper("Поставки.docx");
            for (int i = 0; i < grid1.RowCount - 1; i++)
            {
                var items = new Dictionary<string, string>
                {
                    { "<Name>", Convert.ToString(grid1[1, i].Value) },
                    { "<Buyer>", Convert.ToString(grid1[2, i].Value)},
                    { "<Date>", Convert.ToDateTime(grid1[4, i].Value).Date.ToString()},
                    { "<Count>", Convert.ToString(grid1[3, i].Value) },
                    { "<Sale>", Convert.ToString(grid1[5, i].Value) }
                };
                helper.Process(items);
            }
        }

        public static void SetComboElem(ComboBox comboBox, int id, DataGridView dataGridView, int column)
        {
            for (int i = 0; i < comboBox.Items.Count; i++)
            {
                if (comboBox.Items[i].ToString().Replace(" ", "") == dataGridView[id, column].Value.ToString().Replace(" ", ""))
                {
                    comboBox.SelectedIndex = i;
                    return;
                }
            }
        }

        public static void SetTextBox(TextBox textBox, int id, DataGridView dataGridView, int column)
        {
            textBox.Text = dataGridView[id, column].Value.ToString().Replace(" ", "");
        }

        public static void SetMaskTextBox(MaskedTextBox textBox, int id, DataGridView dataGridView, int column)
        {
            textBox.Text = dataGridView[id, column].Value.ToString().Replace(" ", "");
        }

        public static void textbox(TextBox box)
        {
            string t1 = box.Text;
            if (t1.Length > 0)
            {
                t1 = Convert.ToString(t1[0]).ToUpper();
                for (int i = 1; i < box.Text.Length; i++)
                    t1 += box.Text[i];
                box.Text = t1;

                box.SelectionStart = t1.Length;
            }

        }

        public static bool isYear(TextBox box)
        {
            return (int.Parse(box.Text) >= 0 && int.Parse(box.Text) <= 2022) ? true : false;
        }
    }
}
