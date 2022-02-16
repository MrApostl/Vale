using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Bookstore.BD;
using LiveCharts;
using LiveCharts.Wpf;

namespace Bookstore
{
    public partial class Graph : Form
    {
        public Graph()
        {
            InitializeComponent();
        }

        private void graphh()
        {
            try
            {
                SeriesCollection series = new SeriesCollection();
                ChartValues<int> money = new ChartValues<int>();
                List<string> dates = new List<string>();
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    money.Add(Convert.ToInt32(dataGridView1[6, i].Value));

                    dates.Add(Convert.ToString(dataGridView1[1, i].Value) + "\n" + Convert.ToString(dataGridView1[8, i].Value) + "\n" + Convert.ToDateTime(dataGridView1[9, i].Value).ToShortDateString());
                }
                cartesianChart1.AxisX.Clear();

                cartesianChart1.AxisX.Add(new Axis()
                {
                    Title = "Книга, стоимость и дата заказа",
                    Labels = dates
                });

                LineSeries moneyLine = new LineSeries();

                moneyLine.Title = "Количество";
                moneyLine.Values = money;

                series.Add(moneyLine);
                cartesianChart1.Series = series;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void Graph_Load(object sender, EventArgs e)
        {
            Connection.GetTable("SELECT [Order].[Id],Book.Name as Книга,REPLACE (Buyer.LName, ' ', '' ) + ' ' + REPLACE (Buyer.FName, ' ', '' ) + ' ' + REPLACE (Buyer.Patronymic, ' ', '' ) as Покупатель,REPLACE (Employee.LName, ' ', '' ) + ' ' + REPLACE (Employee.FName, ' ', '' ) + ' ' + REPLACE (Employee.Patronymic, ' ', '' ) as Сотрудник,Discount.Amount as [Скидка, %], (Discount.Amount * Book.Sale / 100) as Скидка,[Count] as Количество, Book.Sale as Цена,Book.Sale - (Discount.Amount * Book.Sale / 100) as [Конечная цена],[DateOrder] as [Дата заказа]  FROM (((Book inner join [Order] on Book.Id = [Order].[Id_book]) inner join Employee on Employee.Id = [Order].Id_employee) inner join Discount on Discount.Id = [Order].Id_discount) inner join Buyer on Buyer.Id = [Order].Id_buyer", dataGridView1);
            //cartesianChart1.LegendLocation = LegendLocation.Bottom;
            graphh();
        }
    }
}
