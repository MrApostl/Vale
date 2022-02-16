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
    public partial class Graph2 : Form
    {
        public Graph2()
        {
            InitializeComponent();
        }

        private void Graph2_Load(object sender, EventArgs e)
        {
            Connection.GetTable("select sum(Book.Sale - (Discount.Amount * Book.Sale / 100)) as [Конечная цена], REPLACE (Employee.LName, ' ', '' ) + ' ' + REPLACE (Employee.FName, ' ', '' ) + ' ' + REPLACE (Employee.Patronymic, ' ', '' ) as Сотрудник FROM (((Book inner join [Order] on Book.Id = [Order].[Id_book]) inner join Employee on Employee.Id = [Order].Id_employee) inner join Discount on Discount.Id = [Order].Id_discount) inner join Buyer on Buyer.Id = [Order].Id_buyer  group by REPLACE (Employee.LName, ' ', '' ) + ' ' + REPLACE (Employee.FName, ' ', '' ) + ' ' + REPLACE (Employee.Patronymic, ' ', '' )", dataGridView1);
            //cartesianChart1.LegendLocation = LegendLocation.Bottom;
            graphh();
        }
        private void graphh()
        {
            try
            {
                SeriesCollection series = new SeriesCollection();
                ChartValues<double> money = new ChartValues<double>();
                List<string> dates = new List<string>();
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    money.Add(Convert.ToDouble(dataGridView1[0, i].Value));

                    dates.Add(Convert.ToString(dataGridView1[1, i].Value));
                }
                cartesianChart1.AxisX.Clear();

                cartesianChart1.AxisX.Add(new Axis()
                {
                    Title = "Cотрудник",
                    Labels = dates
                });

                LineSeries moneyLine = new LineSeries();

                moneyLine.Title = "Сумма";
                moneyLine.Values = money;

                series.Add(moneyLine);
                cartesianChart1.Series = series;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
