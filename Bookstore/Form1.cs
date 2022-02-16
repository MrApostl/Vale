using Bookstore.BD;
using Bookstore.Framework;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Bookstore
{
    public partial class Form1 : Form
    {
        private bool isRed = false;

        public Form1()
        {
            InitializeComponent();
            SetTables();
        }

        public void SetTables()
        {
            Connection.GetTable("SELECT Book.[Id],Genre.Name as Жанр, REPLACE (Author.LName, ' ', '' ) + ' ' + REPLACE (Author.FName, ' ', '' )+ ' ' + REPLACE (Author.Patronymic, ' ', '' ) as Автор,Publishing.Name as Издательство,Language.Name as Язык,Book.[Name] as Название,[DateWriting] as [Дата написания],[Sale] as Цена FROM ((((Genre inner join Book on Genre.Id = Book.Id_genre) inner join Author on Author.Id = Book.Id_author) inner join Publishing on Publishing.Id = Book.Id_publishing) inner join Language on Language.Id = Book.Id_language)", dataGridView1);
            Connection.GetTable("Select Book.Id, Book.Name as Книга, ISNULL([Entrance].Count, 0)  as [На складе], ISNULL([Order].Count, 0) as Заказано From (Book left join [Entrance] on Book.Id = Entrance.Id_book) left join [Order] on  Book.Id = [Order].Id_book group by Book.Id, Book.Name, [Entrance].Count, [Order].Count", dataGridView2);
            Connection.GetTable("SELECT [Id],[Name] as Название,[Adress] as Сайт,[Telephone] as Телефон FROM [Publishing]", dataGridView3);
            Connection.GetTable("SELECT [Id],[Name] as Название FROM [Genre]", dataGridView4);
            Connection.GetTable("SELECT [Id],[LName] as Фамилия,[FName] as Имя,[Patronymic] as Отчество FROM [Author]", dataGridView5);
            Connection.GetTable("SELECT [Id],[Name] as Название FROM [bookstoreBD].[dbo].[Language]", dataGridView6);
            Connection.GetTable("SELECT [Id],[LName] as Фамилия,[FName] as Имя,[Patronymic] as Отчество,[Telephone] as Телефон,[Email] as Почта FROM [Buyer]", dataGridView7);
            Connection.GetTable("SELECT [Employee].Id,[LName] as Фамилия,[FName] as Имя,[Patronymic] as Отчество,Post.Name as Должность FROM Post inner join [Employee] on Post.Id = Employee.Id_post", dataGridView8);
            Connection.GetTable("SELECT [Id],[Name] as Название FROM [bookstoreBD].[dbo].[Post]", dataGridView9);
            Connection.GetTable("SELECT Entrance.[Id],Book.Name as Книга,Supplier.Name as Поставщик,[Count] as Количество,[Date] as [Дата поставки], Book.Sale * Count as Цена FROM (Book inner join [Entrance] on Book.Id = Entrance.Id_book) inner join Supplier on Supplier.Id = Entrance.Id_supplier", dataGridView10);
            Connection.GetTable("SELECT [Order].[Id],Book.Name as Книга,REPLACE (Buyer.LName, ' ', '' ) + ' ' + REPLACE (Buyer.FName, ' ', '' ) + ' ' + REPLACE (Buyer.Patronymic, ' ', '' ) as Покупатель,REPLACE (Employee.LName, ' ', '' ) + ' ' + REPLACE (Employee.FName, ' ', '' ) + ' ' + REPLACE (Employee.Patronymic, ' ', '' ) as Сотрудник,Discount.Amount as [Скидка, %], (Discount.Amount * Book.Sale / 100) as Скидка,[Count] as Количество, Book.Sale as Цена,Book.Sale - (Discount.Amount * Book.Sale / 100) as [Конечная цена],[DateOrder] as [Дата заказа]  FROM (((Book inner join [Order] on Book.Id = [Order].[Id_book]) inner join Employee on Employee.Id = [Order].Id_employee) inner join Discount on Discount.Id = [Order].Id_discount) inner join Buyer on Buyer.Id = [Order].Id_buyer", dataGridView11);
            Connection.GetTable("SELECT [Id],[Amount] as [Количество, %] FROM [Discount]", dataGridView12);
            Connection.GetTable("SELECT [Id],[Name] as Название,[Telephone] as Телефон,[Adress] as Сайт FROM [bookstoreBD].[dbo].[Supplier]", dataGridView13);
            Utils.addCombo(dataGridView4, comboBox10, 1);
            Utils.addComboAuthor(dataGridView5, comboBox9);
            Utils.addCombo(dataGridView3, comboBox12, 1);
            Utils.addCombo(dataGridView6, comboBox11, 1);
            Utils.addCombo(dataGridView9, comboBox6, 1);
            Utils.addCombo(dataGridView1, comboBox8, 5);
            Utils.addCombo(dataGridView13, comboBox7, 1);
            Utils.addCombo(dataGridView1, comboBox1, 5);
            Utils.addComboBuyer(dataGridView7, comboBox2); 
            Utils.addComboEmployee(dataGridView8, comboBox3);
            Utils.addCombo(dataGridView12, comboBox4, 1);
            isRed = false;
        }
       
        private void SetSupplyForPeriod(DataGridView dataGridView, MaskedTextBox timePicker1, MaskedTextBox timePicker2)
        {
            Connection.GetTable($"SELECT Entrance.[Id],Book.Name as Книга,Supplier.Name as Поставщик,[Count] as Количество,[Date] as [Дата поставки], Book.Sale * Count as Цена FROM (Book inner join [Entrance] on Book.Id = Entrance.Id_book) inner join Supplier on Supplier.Id = Entrance.Id_supplier where [Date] BETWEEN CONVERT(date,'{Convert.ToDateTime(timePicker1.Text).ToShortDateString()}',104) and CONVERT(date,'{Convert.ToDateTime(timePicker2.Text).ToShortDateString()}',104)", dataGridView);
        }

        private void SetSumForPeriod(DataGridView dataGridView, MaskedTextBox timePicker1, MaskedTextBox timePicker2)
        {
            Connection.GetTable($"select [Order].[Id], [DateOrder] as [Дата заказа], Sum(Book.Sale - (Discount.Amount * Book.Sale / 100)) as [Конечная цена]  FROM (((Book inner join [Order] on Book.Id = [Order].[Id_book]) inner join Employee on Employee.Id = [Order].Id_employee) inner join Discount on Discount.Id = [Order].Id_discount) inner join Buyer on Buyer.Id = [Order].Id_buyer where [DateOrder] BETWEEN CONVERT(date,'{Convert.ToDateTime(timePicker1.Text).ToShortDateString()}',104) and CONVERT(date,'{Convert.ToDateTime(timePicker2.Text).ToShortDateString()}',104) Group by [DateOrder], [Order].[Id]", dataGridView);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    Utils.FindRecord(dataGridView1, textBox28);
                    break;
                case 1:
                    Utils.FindRecord(dataGridView2, textBox28);
                    break;
                case 2:
                    Utils.FindRecord(dataGridView3, textBox28);
                    break;
                case 3:
                    Utils.FindRecord(dataGridView4, textBox28);
                    break;
                case 4:
                    Utils.FindRecord(dataGridView5, textBox28);
                    break;
                case 5:
                    Utils.FindRecord(dataGridView6, textBox28);
                    break;
                case 6:
                    Utils.FindRecord(dataGridView7, textBox28);
                    break;
                case 7:
                    Utils.FindRecord(dataGridView8, textBox28);
                    break;
                case 8:
                    Utils.FindRecord(dataGridView9, textBox28);
                    break;
                case 9:
                    Utils.FindRecord(dataGridView13, textBox28);
                    break;
                case 10:
                    Utils.FindRecord(dataGridView10, textBox28);
                    break;
                case 11:
                    Utils.FindRecord(dataGridView11, textBox28);
                    break;
                case 12:
                    Utils.FindRecord(dataGridView12, textBox28);
                    break;
                default:
                    break;
            }
            textBox28.Text = "";
        }

        private void button20_Click(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    Utils.FilterGrid(dataGridView1, comboBox14, textBox29);
                    break;
                case 1:
                    Utils.FilterGrid(dataGridView2, comboBox14, textBox29);
                    break;
                case 2:
                    Utils.FilterGrid(dataGridView3, comboBox14, textBox29);
                    break;
                case 3:
                    Utils.FilterGrid(dataGridView4, comboBox14, textBox29);
                    break;
                case 4:
                    Utils.FilterGrid(dataGridView5, comboBox14, textBox29);
                    break;
                case 5:
                    Utils.FilterGrid(dataGridView6, comboBox14, textBox29);
                    break;
                case 6:
                    Utils.FilterGrid(dataGridView7, comboBox14, textBox29);
                    break;
                case 7:
                    Utils.FilterGrid(dataGridView8, comboBox14, textBox29);
                    break;
                case 8:
                    Utils.FilterGrid(dataGridView9, comboBox14, textBox29);
                    break;
                case 9:
                    Utils.FilterGrid(dataGridView13, comboBox14, textBox29);
                    break;
                case 10:
                    Utils.FilterGrid(dataGridView10, comboBox14, textBox29);
                    break;
                case 11:
                    Utils.FilterGrid(dataGridView11, comboBox14, textBox29);
                    break;
                case 12:
                    Utils.FilterGrid(dataGridView12, comboBox14, textBox29);
                    break;
                default:
                    break;
            }
            textBox29.Text = "";
            comboBox14.SelectedIndex = -1;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            panel15.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (Utils.isTextBoxNull(textBox8))
            {
                if (!isRed)
                    Connection.AddRecord($"INSERT INTO [Post](Name) VALUES ('{textBox8.Text}')");
                else Connection.AddRecord($"Update [Post] set Name = '{textBox8.Text}' where Id = '{dataGridView9[0, dataGridView9.CurrentCell.RowIndex].Value}'");
                textBox8.Text = "";
                button5.Text = "Добавить";
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Utils.isTextBoxNull(textBox1))
            {
                if (!isRed)
                    Connection.AddRecord($"INSERT INTO [Discount](Amount) VALUES ('{textBox1.Text}')");
                else Connection.AddRecord($"Update [Discount] set Amount = '{textBox1.Text}' where Id = '{dataGridView12[0, dataGridView12.CurrentCell.RowIndex].Value}'");
                textBox1.Text = "";
                button1.Text = "Добавить";
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (Utils.isTextBoxNull(textBox17))
            {
                if (!isRed)
                    Connection.AddRecord($"INSERT INTO [Language](Name) VALUES ('{textBox17.Text}')");
                else Connection.AddRecord($"Update [Language] set Name = '{textBox17.Text}' where Id = '{dataGridView6[0, dataGridView6.CurrentCell.RowIndex].Value}'");
                textBox17.Text = "";
                button8.Text = "Добавить";
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (Utils.isTextBoxNull(textBox21))
            {
                if (!isRed)
                    Connection.AddRecord($"INSERT INTO [Genre](Name) VALUES ('{textBox21.Text}')");
                else Connection.AddRecord($"Update [Genre] set Name = '{textBox21.Text}' where Id = '{dataGridView4[0, dataGridView4.CurrentCell.RowIndex].Value}'");
                textBox21.Text = "";
                button10.Text = "Добавить";
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if(Utils.isTextBoxNull(textBox24) && Utils.isTextBoxNull(textBox23) && Utils.isTextBoxNull(textBox22))
            {
                if (!isRed)
                    Connection.AddRecord($"INSERT INTO [Publishing](Name, Adress, Telephone) VALUES ('{textBox24.Text}', '{textBox23.Text}', '{textBox22.Text}')");
                else Connection.AddRecord($"Update [Publishing] set Name = '{textBox24.Text}', Adress = '{textBox23.Text}', Telephone = '{textBox22.Text}' where Id = '{dataGridView3[0, dataGridView3.CurrentCell.RowIndex].Value}'");
                textBox24.Text = "";
                textBox23.Text = "";
                textBox22.Text = "";
                button11.Text = "Добавить";
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (Utils.isTextBoxNull(textBox7) && Utils.isTextBoxNull(textBox5) && Utils.isTextBoxNull(textBox6))
            {
                if (!isRed)
                    Connection.AddRecord($"INSERT INTO [Supplier](Name, Telephone, Adress) VALUES ('{textBox7.Text}', '{textBox5.Text}', '{textBox6.Text}')");
                else Connection.AddRecord($"Update [Supplier] set Name = '{textBox7.Text}', Adress = '{textBox6.Text}', Telephone = '{textBox5.Text}' where Id = '{dataGridView13[0, dataGridView13.CurrentCell.RowIndex].Value}'");
                textBox7.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                button4.Text = "Добавить";
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (Utils.isTextBoxNull(textBox10) && Utils.isTextBoxNull(textBox13) && Utils.isTextBoxNull(textBox14) && Utils.isTextBoxNull(textBox15) && Utils.isTextBoxNull(textBox16))
            {
                if (!isRed)
                    Connection.AddRecord($"INSERT INTO [Buyer](LName, FName, Patronymic,Telephone, Email) VALUES ('{textBox10.Text}', '{textBox13.Text}', '{textBox14.Text}', '{textBox15.Text}', '{textBox16.Text}')");
                else Connection.AddRecord($"Update [Buyer] set LName = '{textBox10.Text}', FName = '{textBox13.Text}', Patronymic = '{textBox14.Text}', Telephone = '{textBox15.Text}', Email = '{textBox16.Text}' where Id = '{dataGridView7[0, dataGridView7.CurrentCell.RowIndex].Value}'");
                textBox10.Text = "";
                textBox13.Text = "";
                textBox14.Text = "";
                textBox15.Text = "";
                textBox16.Text = "";
                button7.Text = "Добавить";
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (Utils.isTextBoxNull(textBox24) && Utils.isTextBoxNull(textBox23) && Utils.isTextBoxNull(textBox22))
            {
                if (!isRed)
                    Connection.AddRecord($"INSERT INTO [Author](LName, FName, Patronymic) VALUES ('{textBox20.Text}', '{textBox19.Text}', '{textBox18.Text}')");
                else Connection.AddRecord($"Update [Author] set LName = '{textBox20.Text}', FName = '{textBox19.Text}', Patronymic = '{textBox18.Text}' where Id = '{dataGridView5[0, dataGridView5.CurrentCell.RowIndex].Value}'");
                textBox20.Text = "";
                textBox19.Text = "";
                textBox18.Text = "";
                button9.Text = "Добавить";
            }
            else 
            {
                MessageBox.Show("Не все поля заполнены");
            } 
        }

        private void button13_Click(object sender, EventArgs e)
        {
            int id_genre = 0;
            int id_author = 0;
            int id_publishing = 0;
            int id_language = 0;

            if(Utils.isTextBoxNull(textBox30) && Utils.isTextBoxNull(textBox26) && Utils.isComboBoxNull(comboBox10) && Utils.isComboBoxNull(comboBox9) && Utils.isComboBoxNull(comboBox12) && Utils.isComboBoxNull(comboBox11) && Utils.isMaskTextBoxNull(maskedTextBox2))
            {
                for (int i = 0; i < dataGridView4.Rows.Count - 1; i++)
                {
                    if(comboBox10.Items[comboBox10.SelectedIndex].ToString().Replace(" ", "") == dataGridView4[1, i].Value.ToString().Replace(" ", ""))
                    {
                        id_genre = Convert.ToInt32(dataGridView4[0, i].Value);
                        for (int j = 0; j < dataGridView5.Rows.Count - 1; j++)
                        {
                            if (comboBox9.Items[comboBox9.SelectedIndex].ToString().Replace(" ", "") == (dataGridView5[1, j].Value + " " + dataGridView5[2, j].Value + " " + dataGridView5[3, j].Value).Replace(" ", ""))
                            {
                                id_author = Convert.ToInt32(dataGridView5[0, j].Value);
                                for (int z = 0; z < dataGridView3.Rows.Count - 1; z++)
                                {
                                    if (comboBox12.Items[comboBox12.SelectedIndex].ToString().Replace(" ", "") == dataGridView3[1, z].Value.ToString().Replace(" ", ""))
                                    {
                                        id_publishing = Convert.ToInt32(dataGridView3[0, z].Value);
                                        for (int x = 0; x < dataGridView6.Rows.Count - 1; x++)
                                        {
                                            if (comboBox11.Items[comboBox11.SelectedIndex].ToString().Replace(" ", "") == dataGridView6[1, x].Value.ToString().Replace(" ", ""))
                                            {
                                                try
                                                {
                                                    id_language = Convert.ToInt32(dataGridView3[0, x].Value);
                                                    if (!isRed)
                                                        Connection.AddRecord($"INSERT INTO [Book](Id_genre, Id_author, Id_publishing,Id_language, Name, DateWriting, Sale) VALUES ('{id_genre}', '{id_author}', '{id_publishing}', '{id_language}', '{textBox30.Text}', '{maskedTextBox2.Text}', '{textBox26.Text}')");
                                                    else Connection.AddRecord($"update [Book] set Id_genre = '{id_genre}', Id_author ='{id_author}', Id_publishing = '{id_publishing}',Id_language = '{id_language}', Name = '{textBox30.Text}', DateWriting = '{maskedTextBox2.Text}', Sale ='{textBox26.Text}' where Id = '{dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value}'");
                                                }
                                                catch
                                                {
                                                    MessageBox.Show("Не верно введена дата. Пример: месяц.день.год");
                                                    Connection.CloseConnection();
                                                }
                                                comboBox10.SelectedIndex = -1;
                                                comboBox9.SelectedIndex = -1;
                                                comboBox12.SelectedIndex = -1;
                                                comboBox11.SelectedIndex = -1;
                                                textBox30.Text = "";
                                                maskedTextBox2.Text = "";
                                                textBox26.Text = "";
                                                button13.Text = "Добавить";
                                                return;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int id_post = 0;
            if (Utils.isTextBoxNull(textBox9) && Utils.isTextBoxNull(textBox11) && Utils.isTextBoxNull(textBox12) && Utils.isComboBoxNull(comboBox6))
            {
                for (int x = 0; x < dataGridView9.Rows.Count - 1; x++)
                {
                    if (comboBox6.Items[comboBox6.SelectedIndex].ToString().Replace(" ", "") == dataGridView9[1, x].Value.ToString().Replace(" ", ""))
                    {
                        id_post = Convert.ToInt32(dataGridView9[0, x].Value);
                        if (!isRed)
                            Connection.AddRecord($"INSERT INTO [Employee](LName, FName, Patronymic, Id_post) VALUES ('{textBox9.Text}', '{textBox11.Text}', '{textBox12.Text}', '{id_post}')");
                        else Connection.AddRecord($"Update [Employee] set LName = '{textBox9.Text}', FName = '{textBox11.Text}', Patronymic = '{textBox12.Text}', Id_post = '{id_post}' where Id = '{dataGridView8[0, dataGridView8.CurrentCell.RowIndex].Value}'");
                        comboBox6.SelectedIndex = -1;
                        textBox9.Text = "";
                        textBox11.Text = "";
                        textBox12.Text = "";
                        button6.Text = "Добавить";
                        return;
                    }
                }
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int id_book = 0;
            int id_supplier = 0;
            if (Utils.isTextBoxNull(textBox4) && Utils.isMaskTextBoxNull(maskedTextBox1) && Utils.isComboBoxNull(comboBox7) && Utils.isComboBoxNull(comboBox8))
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    if (comboBox8.Items[comboBox8.SelectedIndex].ToString().Replace(" ", "") == dataGridView1[5, i].Value.ToString().Replace(" ", ""))
                    {
                        id_book = Convert.ToInt32(dataGridView1[0, i].Value);
                        for (int j = 0; j < dataGridView13.Rows.Count - 1; j++)
                        {
                            if (comboBox7.Items[comboBox7.SelectedIndex].ToString().Replace(" ", "") == (dataGridView13[1, j].Value).ToString().Replace(" ", ""))
                            {
                                try
                                {
                                    id_supplier = Convert.ToInt32(dataGridView13[0, j].Value);
                                    if (!isRed)
                                        Connection.AddRecord($"INSERT INTO [Entrance](Id_book, Id_supplier, Count, Date) VALUES ('{id_book}', '{id_supplier}', '{textBox4.Text}', '{maskedTextBox1.Text}')");
                                    else Connection.AddRecord($"Update [Entrance] set Id_book = '{id_book}', Id_supplier = '{id_supplier}', Count = '{textBox4.Text}', Date = '{maskedTextBox1.Text}' where Id = '{dataGridView10[0, dataGridView10.CurrentCell.RowIndex].Value}'");
                                }
                                catch
                                {
                                    MessageBox.Show("Не верно введена дата. Пример: месяц.день.год");
                                    Connection.CloseConnection();
                                }
                                comboBox8.SelectedIndex = -1;
                                comboBox7.SelectedIndex = -1;
                                textBox4.Text = "";
                                maskedTextBox1.Text = "";
                                button3.Text = "Добавить";
                                return;
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int id_book = 0;
            int id_buyer = 0;
            int id_employee = 0;
            int id_discount = 0;
            if (Utils.isTextBoxNull(textBox2) && Utils.isComboBoxNull(comboBox1) && Utils.isComboBoxNull(comboBox2) && Utils.isComboBoxNull(comboBox3) && Utils.isComboBoxNull(comboBox4))
            {
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    if (comboBox1.Items[comboBox1.SelectedIndex].ToString().Replace(" ", "") == dataGridView1[5, i].Value.ToString().Replace(" ", ""))
                    {
                        id_book = Convert.ToInt32(dataGridView4[0, i].Value);
                        for (int j = 0; j < dataGridView7.Rows.Count - 1; j++)
                        {
                            if (comboBox2.Items[comboBox2.SelectedIndex].ToString().Replace(" ", "") == (dataGridView7[1, j].Value + " " + dataGridView7[2, j].Value + " " + dataGridView7[3, j].Value).Replace(" ", ""))
                            {
                                id_buyer = Convert.ToInt32(dataGridView7[0, j].Value);
                                for (int z = 0; z < dataGridView8.Rows.Count - 1; z++)
                                {
                                    if (comboBox3.Items[comboBox3.SelectedIndex].ToString().Replace(" ", "") == (dataGridView8[1, z].Value + " " + dataGridView8[2, z].Value + " " + dataGridView8[3, z].Value).Replace(" ", ""))
                                    {
                                        id_employee = Convert.ToInt32(dataGridView8[0, z].Value);
                                        for (int x = 0; x < dataGridView12.Rows.Count - 1; x++)
                                        {
                                            if (comboBox4.Items[comboBox4.SelectedIndex].ToString().Replace(" ", "") == dataGridView12[1, x].Value.ToString().Replace(" ", ""))
                                            {
                                                try
                                                {
                                                    id_discount = Convert.ToInt32(dataGridView12[0, x].Value);
                                                    if (!isRed)
                                                        Connection.AddRecord($"INSERT INTO [Order](id_book, id_buyer, id_employee,id_discount, Count, DateOrder) VALUES ('{id_book}', '{id_buyer}', '{id_employee}', '{id_discount}', '{textBox2.Text}', '{maskedTextBox3.Text}')");
                                                    else Connection.AddRecord($"Update [Order] set id_book ='{id_book}', id_buyer = '{id_buyer}', id_employee = '{id_employee}', id_discount = '{id_discount}', Count = '{textBox2.Text}' , DateOrder = '{maskedTextBox3.Text}' where Id = '{dataGridView11[0, dataGridView11.CurrentCell.RowIndex].Value}'");
                                                }
                                                catch (Exception)
                                                {
                                                    MessageBox.Show("Не верно введена дата. Пример: месяц.день.год");
                                                    Connection.CloseConnection();
                                                }
                                                
                                                comboBox1.SelectedIndex = -1;
                                                comboBox2.SelectedIndex = -1;
                                                comboBox3.SelectedIndex = -1;
                                                comboBox4.SelectedIndex = -1;
                                                textBox2.Text = "";
                                                button2.Text = "Добавить";
                                                maskedTextBox3.Text = "";
                                                return;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                comboBox1.SelectedIndex = -1;
                comboBox2.SelectedIndex = -1;
                comboBox3.SelectedIndex = -1;
                comboBox4.SelectedIndex = -1;
                textBox2.Text = "";
                button2.Text = "Добавить";
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }
                

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            button13.Text = "Редактировать";
            Utils.SetComboElem(comboBox10, 1, dataGridView1, dataGridView1.CurrentCell.RowIndex);
            Utils.SetComboElem(comboBox9, 2, dataGridView1, dataGridView1.CurrentCell.RowIndex);
            Utils.SetComboElem(comboBox12, 3, dataGridView1, dataGridView1.CurrentCell.RowIndex);
            Utils.SetComboElem(comboBox11, 4, dataGridView1, dataGridView1.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox30, 5, dataGridView1, dataGridView1.CurrentCell.RowIndex);
            Utils.SetMaskTextBox(maskedTextBox2, 6, dataGridView1, dataGridView1.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox26, 7, dataGridView1, dataGridView1.CurrentCell.RowIndex);
            isRed = true;
        }

        private void ClearText()
        {
            isRed = false;
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            textBox2.Text = "";
            button2.Text = "Добавить";
            comboBox8.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            textBox4.Text = "";
            maskedTextBox1.Text = "";
            button3.Text = "Добавить";
            comboBox6.SelectedIndex = -1;
            textBox9.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            button6.Text = "Добавить";
            comboBox10.SelectedIndex = -1;
            comboBox9.SelectedIndex = -1;
            comboBox12.SelectedIndex = -1;
            comboBox11.SelectedIndex = -1;
            textBox30.Text = "";
            maskedTextBox2.Text = "";
            textBox26.Text = "";
            button13.Text = "Добавить";
            textBox20.Text = "";
            textBox19.Text = "";
            textBox18.Text = "";
            button9.Text = "Добавить";
            textBox10.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox16.Text = "";
            button7.Text = "Добавить";
            textBox7.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            button4.Text = "Добавить";
            textBox24.Text = "";
            textBox23.Text = "";
            textBox22.Text = "";
            button11.Text = "Добавить";
            textBox8.Text = "";
            button5.Text = "Добавить";
            textBox1.Text = "";
            button1.Text = "Добавить";
            textBox17.Text = "";
            button8.Text = "Добавить";
            textBox21.Text = "";
            button10.Text = "Добавить";
            textBox8.Text = "";
            button5.Text = "Добавить";
            maskedTextBox3.Text = "";
            maskedTextBox4.Text = "";
            maskedTextBox5.Text = "";
            maskedTextBox6.Text = "";
            maskedTextBox7.Text = "";
        }

        private void dataGridView3_Click(object sender, EventArgs e)
        {
            button11.Text = "Редактировать";
            Utils.SetTextBox(textBox24, 1, dataGridView3, dataGridView3.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox23, 2, dataGridView3, dataGridView3.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox22, 3, dataGridView3, dataGridView3.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView4_Click(object sender, EventArgs e)
        {
            button10.Text = "Редактировать";
            Utils.SetTextBox(textBox21, 1, dataGridView4, dataGridView4.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView5_Click(object sender, EventArgs e)
        {
            button9.Text = "Редактировать";
            Utils.SetTextBox(textBox20, 1, dataGridView5, dataGridView5.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox19, 2, dataGridView5, dataGridView5.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox18, 3, dataGridView5, dataGridView5.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView6_Click(object sender, EventArgs e)
        {
            button8.Text = "Редактировать";
            Utils.SetTextBox(textBox17, 1, dataGridView6, dataGridView6.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView7_Click(object sender, EventArgs e)
        {
            button7.Text = "Редактировать";
            Utils.SetTextBox(textBox10, 1, dataGridView7, dataGridView7.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox13, 2, dataGridView7, dataGridView7.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox14, 3, dataGridView7, dataGridView7.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox15, 4, dataGridView7, dataGridView7.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox16, 5, dataGridView7, dataGridView7.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView8_Click(object sender, EventArgs e)
        {
            button6.Text = "Редактировать";
            Utils.SetTextBox(textBox9, 1, dataGridView8, dataGridView8.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox11, 2, dataGridView8, dataGridView8.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox12, 3, dataGridView8, dataGridView8.CurrentCell.RowIndex);
            Utils.SetComboElem(comboBox6, 4, dataGridView8, dataGridView8.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView9_Click(object sender, EventArgs e)
        {
            button5.Text = "Редактировать";
            Utils.SetTextBox(textBox8, 1, dataGridView9, dataGridView9.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView13_Click(object sender, EventArgs e)
        {
            button4.Text = "Редактировать";
            Utils.SetTextBox(textBox7, 1, dataGridView13, dataGridView13.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox5, 2, dataGridView13, dataGridView13.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox6, 3, dataGridView13, dataGridView13.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView10_Click(object sender, EventArgs e)
        {
            button3.Text = "Редактировать";
            Utils.SetComboElem(comboBox8, 1, dataGridView10, dataGridView10.CurrentCell.RowIndex);
            Utils.SetComboElem(comboBox7, 2, dataGridView10, dataGridView10.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox4, 3, dataGridView10, dataGridView10.CurrentCell.RowIndex);
            Utils.SetMaskTextBox(maskedTextBox1, 4, dataGridView10, dataGridView10.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView11_Click(object sender, EventArgs e)
        {
            button2.Text = "Редактировать";
            Utils.SetComboElem(comboBox1, 1, dataGridView11, dataGridView11.CurrentCell.RowIndex);
            Utils.SetComboElem(comboBox2, 2, dataGridView11, dataGridView11.CurrentCell.RowIndex);
            Utils.SetComboElem(comboBox3, 3, dataGridView11, dataGridView11.CurrentCell.RowIndex);
            Utils.SetComboElem(comboBox4, 4, dataGridView11, dataGridView11.CurrentCell.RowIndex);
            Utils.SetTextBox(textBox2, 6, dataGridView11, dataGridView11.CurrentCell.RowIndex);
            Utils.SetMaskTextBox(maskedTextBox3, 9, dataGridView11, dataGridView11.CurrentCell.RowIndex);
            isRed = true;
        }

        private void dataGridView12_Click(object sender, EventArgs e)
        {
            button1.Text = "Редактировать";
            Utils.SetTextBox(textBox1, 1, dataGridView12, dataGridView12.CurrentCell.RowIndex);
            isRed = true;
        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox30);
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox24);
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox21);
        }

        private void textBox20_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox20);
        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox19);
        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox18);
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox10);
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox13);
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox14);
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox9);
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox11);
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox12);
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox8);
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            Utils.textbox(textBox7);
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | (Char.IsControl(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | (Char.IsControl(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | (Char.IsControl(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | (Char.IsControl(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | (Char.IsControl(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox22_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | (Char.IsControl(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox26_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) | (Char.IsControl(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox30_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox19_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox18_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox17_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && (!Char.IsPunctuation(e.KeyChar)) && !(e.KeyChar == (char)Keys.Space)) return;
            else
                e.Handled = true;
        }

        private void maskedTextBox2_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
            
        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
            
        }

        private void button19_Click(object sender, EventArgs e)
        {
            
        }

        private void ОбновитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SetTables();
            ClearText();
        }

        private void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dialogResult = new DialogResult();
                switch (tabControl1.SelectedIndex)
                {
                    case 0:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView1.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Book] where Id = '{dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value}'");
                        break;
                    case 2:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView3.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Publishing] where Id = '{dataGridView3[0, dataGridView3.CurrentCell.RowIndex].Value}'");
                        break;
                    case 3:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView4.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Genre] where Id = '{dataGridView4[0, dataGridView4.CurrentCell.RowIndex].Value}'");
                        break;
                    case 4:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView5.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Author] where Id = '{dataGridView5[0, dataGridView5.CurrentCell.RowIndex].Value}'");
                        break;
                    case 5:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView6.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Language] where Id = '{dataGridView6[0, dataGridView6.CurrentCell.RowIndex].Value}'");
                        break;
                    case 6:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView7.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Buyer] where Id = '{dataGridView7[0, dataGridView7.CurrentCell.RowIndex].Value}'");
                        break;
                    case 7:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView8.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Employee] where Id = '{dataGridView8[0, dataGridView8.CurrentCell.RowIndex].Value}'");
                        break;
                    case 8:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView9.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Post] where Id = '{dataGridView9[0, dataGridView9.CurrentCell.RowIndex].Value}'");
                        break;
                    case 9:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView13.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Supplier] where Id = '{dataGridView13[0, dataGridView13.CurrentCell.RowIndex].Value}'");
                        break;
                    case 10:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView10.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Entrance] where Id = '{dataGridView10[0, dataGridView10.CurrentCell.RowIndex].Value}'");
                        break;
                    case 11:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView11.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Order] where Id = '{dataGridView11[0, dataGridView11.CurrentCell.RowIndex].Value}'");
                        break;
                    case 12:
                        dialogResult = MessageBox.Show($"Вы действительно хотите удалить строку № {dataGridView12.CurrentRow.Index + 1}", "Подтверждение удаления", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                            Connection.AddRecord($"Delete From [Dicount] where Id = '{dataGridView12[0, dataGridView12.CurrentCell.RowIndex].Value}'");
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {
                MessageBox.Show($"Данная запись используется в других таблицах, для её удаления необходимо изменить соответствующие записи");
            }
        }

        private void очиститьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ClearText();
        }

        private void фильтрацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel15.Visible = true;
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    Utils.ComboGrid(dataGridView1, comboBox14);
                    break;
                case 1:
                    Utils.ComboGrid(dataGridView2, comboBox14);
                    break;
                case 2:
                    Utils.ComboGrid(dataGridView3, comboBox14);
                    break;
                case 3:
                    Utils.ComboGrid(dataGridView4, comboBox14);
                    break;
                case 4:
                    Utils.ComboGrid(dataGridView5, comboBox14);
                    break;
                case 5:
                    Utils.ComboGrid(dataGridView6, comboBox14);
                    break;
                case 6:
                    Utils.ComboGrid(dataGridView7, comboBox14);
                    break;
                case 7:
                    Utils.ComboGrid(dataGridView8, comboBox14);
                    break;
                case 8:
                    Utils.ComboGrid(dataGridView9, comboBox14);
                    break;
                case 9:
                    Utils.ComboGrid(dataGridView13, comboBox14);
                    break;
                case 10:
                    Utils.ComboGrid(dataGridView10, comboBox14);
                    break;
                case 11:
                    Utils.ComboGrid(dataGridView11, comboBox14);
                    break;
                case 12:
                    Utils.ComboGrid(dataGridView12, comboBox14);
                    break;
                default:
                    break;
            }
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    try
                    {
                        Utils.SaveTable(dataGridView1, "Книги");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 1:
                    try
                    {
                        Utils.SaveTable(dataGridView2, "Хранилище");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 2:
                    try
                    {
                        Utils.SaveTable(dataGridView3, "Издательства");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 3:
                    try
                    {
                        Utils.SaveTable(dataGridView4, "Жанры");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 4:
                    try
                    {
                        Utils.SaveTable(dataGridView5, "Авторы");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 5:
                    try
                    {
                        Utils.SaveTable(dataGridView6, "Языки");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 6:
                    try
                    {
                        Utils.SaveTable(dataGridView7, "Покупатели");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 7:
                    try
                    {
                        Utils.SaveTable(dataGridView8, "Сотрудники");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 8:
                    try
                    {
                        Utils.SaveTable(dataGridView9, "Должности");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 9:
                    try
                    {
                        Utils.SaveTable(dataGridView13, "Поставщики");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 10:
                    try
                    {
                        Utils.SaveTable(dataGridView10, "Поставки");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 11:
                    try
                    {
                        Utils.SaveTable(dataGridView11, "Заказы");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                case 12:
                    try
                    {
                        Utils.SaveTable(dataGridView12, "Скидки");
                    }
                    catch (Exception)
                    { MessageBox.Show($"Таблица не сохранена в Excel"); }
                    break;
                default:
                    break;
            }

        }

        private void книжныйПаспортToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.WordBook(dataGridView1, dataGridView2);
        }

        private void заказыToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Utils.WordOrder(dataGridView11);
        }

        private void статистикаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void книгиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 0);
        }

        private void хранилищеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 1);
        }

        private void издательстваToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 2);
        }

        private void жанрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 3);
        }

        private void авторыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 4);
        }

        private void языкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 5);
        }

        private void покупателиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 6);
        }

        private void сотрудникиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 7);
        }

        private void должностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 8);
        }

        private void поставщикиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 9);
        }

        private void поставкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 10);
        }

        private void заказыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 11);
        }

        private void скидкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 12);
        }

        private void поставкиToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Utils.WordSupply(dataGridView10);
        }

        private void купленныеКнигиЗаПериодToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 13);
        }

        private void суммаЗаПериодToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Utils.ChangeTab(tabControl1, 14);
        }

        private void продажиПоСотрудникамЗаПериодToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Graph2 gr = new Graph2();
            gr.Show();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if(Utils.isMaskTextBoxNull(maskedTextBox4) && Utils.isMaskTextBoxNull(maskedTextBox5))
            {
                try
                {
                    SetSupplyForPeriod(dataGridView14, maskedTextBox4, maskedTextBox5);
                }
                catch
                {
                    MessageBox.Show("Не верно введена дата. Пример: месяц.день.год");
                    Connection.CloseConnection();
                }
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (Utils.isMaskTextBoxNull(maskedTextBox7) && Utils.isMaskTextBoxNull(maskedTextBox6))
            {
                try
                {
                    SetSumForPeriod(dataGridView15, maskedTextBox7, maskedTextBox6);
                }
                catch
                {
                    MessageBox.Show("Не верно введена дата. Пример: месяц.день.год");
                    Connection.CloseConnection();
                }
            }
            else
            {
                MessageBox.Show("Не все поля заполнены");
            }
        }

        private void графикПоЗаказамToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Graph gr = new Graph();
            gr.Show();
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 gr = new Form2();
            gr.Show();
        }

        private void maskedTextBox3_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }
    }
    
}
