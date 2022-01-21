using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace la
{
    /// <summary>
    /// Логика взаимодействия для Page1.xaml
    /// </summary>
    public partial class Page1 : Page
    {
        SqlConnection connection;
        public Page1()
        {
            InitializeComponent();
          
        }

        private void Page1_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Page2());
            
        }
        private async void Page1_Loaded(object sender, RoutedEventArgs e)
        {

            string XXX = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\mvideo\Downloads\la\la\Database1.mdf;Integrated Security=True";
            connection = new SqlConnection(XXX);
            await connection.OpenAsync();
            SqlCommand sqlCommand = new SqlCommand("SELECT id, NameofUBI FROM [zxc] WHERE [id] < 151", connection);
            await sqlCommand.ExecuteNonQueryAsync();
            SqlDataAdapter dataAdp = new SqlDataAdapter(sqlCommand);
            DataTable dt = new DataTable("SALAM"); 
            dataAdp.Fill(dt);
            DataGreed.ItemsSource = dt.DefaultView; 
            connection.Close();
        }
    }
}
