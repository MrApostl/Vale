using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Bookstore.BD
{
    public static class Connection
    {
        private static DataSet data;
        private static SqlDataAdapter adapter;
        private static SqlCommand command;
        private static SqlConnection connection = new SqlConnection(@"Data Source=ARTEM\SQLEXPRESS01;Initial Catalog=bookstoreBD;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False");

        private static void OpenConnection()
        {
            connection.Open();
        }

        public static void CloseConnection()
        {
            connection.Close();
        }

        public static void GetTable(string request, DataGridView grid)
        {
            try
            {
                OpenConnection();

                data = new DataSet();
                adapter = new SqlDataAdapter(request, connection);
                adapter.Fill(data);

                grid.DataSource = data.Tables[0];
                grid.Columns[0].Visible = false;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Не удалось провести операцию");
            }
            finally
            {
                CloseConnection();
            }
        }

        public static void AddRecord(string request)
        {
            //try
            //{
                OpenConnection();

                command = new SqlCommand(request, connection);
                command.ExecuteNonQuery();
                CloseConnection();
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Не удалось провести операцию");
            //}
            //finally
            //{
            //    CloseConnection();
            //}

        }
    }
}
