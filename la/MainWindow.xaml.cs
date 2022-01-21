using Microsoft.Win32;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace la
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()


        {

            InitializeComponent();
            MainFrame.Content = new Page1();
        }
        SqlConnection connection;
        int countupdate = 0;

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            string XXX = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\mvideo\Downloads\la\la\Database1.mdf;Integrated Security=True";
            connection = new SqlConnection(XXX);
            await connection.OpenAsync();
            SqlCommand qaz = new SqlCommand("DROP TABLE [zxc]", connection);
            await qaz.ExecuteNonQueryAsync();
            qaz = new SqlCommand("CREATE TABLE zxc " +
                "(id INT IDENTITY NOT NULL PRIMARY KEY, " +
                "[NameofUBI]             NVARCHAR (MAX) NULL, " +
                "[Description]                    NVARCHAR (MAX) NULL, " +
                "[sourceofthethreat]              NVARCHAR (MAX) NULL, " +
                "[TheObjectOfTheThreatImpact]     NVARCHAR (MAX) NULL, " +
                "[violationofconfidentiality] INT            NULL, " +
                "[IntegrityViolation]        INT            NULL, " +
                "[ViolationofAccessibility]        INT            NULL)", connection);
       
            await qaz.ExecuteNonQueryAsync();





            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM zxc ", connection);
            await sqlCommand.ExecuteNonQueryAsync();
            SqlDataAdapter dataAdp = new SqlDataAdapter(sqlCommand);
            DataTable dt = new DataTable("SALAM");
            dataAdp.Fill(dt);
            mygrid.ItemsSource = dt.DefaultView;

        }
        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            asd.Items.Clear();

            try
            {
                SqlCommand sqlCommand = new SqlCommand("DELETE FROM zxc WHERE id=@id", connection);
                sqlCommand.Parameters.AddWithValue("id", q.Text);
                await sqlCommand.ExecuteNonQueryAsync();
                asd.Items.Add("Успешно");
                countupdate++;
                qqq.Items.Clear();
                qqq.Items.Add(countupdate);

            }
            catch (System.Exception ex)
            {

                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                asd.Items.Add("Ошибка");

            }


        }

        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            asd.Items.Clear();

            try
            {
                SqlCommand sqlCommand = new SqlCommand("INSERT INTO zxc (NameofUBI, Description, sourceofthethreat, TheObjectOfTheThreatImpact, " +
                           "violationofconfidentiality, IntegrityViolation, ViolationofAccessibility)VALUES(@NameofUBI, @Description, @sourceofthethreat," +
                           "@TheObjectOfTheThreatImpact, @violationofconfidentiality, @IntegrityViolation, @ViolationofAccessibility)", connection);
                if (!string.IsNullOrEmpty(d.Text) && !string.IsNullOrWhiteSpace(d.Text) &&
                    !string.IsNullOrEmpty(a.Text) && !string.IsNullOrWhiteSpace(a.Text) &&
                    !string.IsNullOrEmpty(z.Text) && !string.IsNullOrWhiteSpace(z.Text) &&
                    !string.IsNullOrEmpty(x.Text) && !string.IsNullOrWhiteSpace(x.Text) &&
                    !string.IsNullOrEmpty(n.Text) && !string.IsNullOrWhiteSpace(n.Text) &&
                    !string.IsNullOrEmpty(m.Text) && !string.IsNullOrWhiteSpace(m.Text) &&
                    !string.IsNullOrEmpty(b.Text) && !string.IsNullOrWhiteSpace(b.Text))
                {
                    sqlCommand.Parameters.AddWithValue("NameofUBI", d.Text);
                    sqlCommand.Parameters.AddWithValue("Description", b.Text);
                    sqlCommand.Parameters.AddWithValue("sourceofthethreat", m.Text);
                    sqlCommand.Parameters.AddWithValue("TheObjectOfTheThreatImpact", n.Text);
                    sqlCommand.Parameters.AddWithValue("violationofconfidentiality", x.Text);
                    sqlCommand.Parameters.AddWithValue("IntegrityViolation", z.Text);
                    sqlCommand.Parameters.AddWithValue("ViolationofAccessibility", a.Text);
                    await sqlCommand.ExecuteNonQueryAsync();
                    asd.Items.Add("Успешно!");
                    countupdate++;
                    qqq.Items.Clear();
                    qqq.Items.Add(countupdate);

                }
            }

            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                asd.Items.Add("Ошибка");
            }

        }

        private async void Button_Click_2(object sender, RoutedEventArgs e)
        {
            asd.Items.Clear();

            try
            {
                if (!string.IsNullOrEmpty(d.Text) && !string.IsNullOrWhiteSpace(d.Text) &&
                    !string.IsNullOrEmpty(a.Text) && !string.IsNullOrWhiteSpace(a.Text) &&
                    !string.IsNullOrEmpty(z.Text) && !string.IsNullOrWhiteSpace(z.Text) &&
                    !string.IsNullOrEmpty(x.Text) && !string.IsNullOrWhiteSpace(x.Text) &&
                    !string.IsNullOrEmpty(n.Text) && !string.IsNullOrWhiteSpace(n.Text) &&
                    !string.IsNullOrEmpty(q.Text) && !string.IsNullOrWhiteSpace(q.Text) &&
                    !string.IsNullOrEmpty(m.Text) && !string.IsNullOrWhiteSpace(m.Text) &&
                    !string.IsNullOrEmpty(b.Text) && !string.IsNullOrWhiteSpace(b.Text))
                {
                    SqlDataReader reader = null;
                    SqlCommand dsa = new SqlCommand("SELECT * FROM zxc", connection);
                    reader = await dsa.ExecuteReaderAsync();
                    try
                    {
                        while (await reader.ReadAsync())

                        {
                            if (q.Text == Convert.ToString(reader["id"]))
                            {
                                www.Items.Add($"БЫЛО:({Convert.ToString(reader["id"])}) ({Convert.ToString(reader["NameofUBI"])}) ({Convert.ToString(reader["Description"])}) " +
                                $"({Convert.ToString(reader["sourceofthethreat"])}) ({Convert.ToString(reader["TheObjectOfTheThreatImpact"])}) ({Convert.ToString(reader["violationofconfidentiality"])}) " +
                                $"({Convert.ToString(reader["IntegrityViolation"])}) ({Convert.ToString(reader["ViolationofAccessibility"])})");
                                www.Items.Add($"Стало:{q.Text}, {d.Text}, {b.Text}, {m.Text}, {n.Text}, {x.Text}, {z.Text}, {a.Text}");
                            }
                        }
                    }
                    catch (Exception ex)

                    {
                        asd.Items.Clear();
                        MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                        asd.Items.Add("Ошибка");
                    }

                    finally
                    {
                        if (reader != null)
                        {
                            reader.Close();
                        }
                    }

                    SqlCommand SS = new SqlCommand("UPDATE zxc SET [NameofUBI] = @NameofUBI, [Description] = @Description, [sourceofthethreat] = @sourceofthethreat, " +
                            "[TheObjectOfTheThreatImpact]=@TheObjectOfTheThreatImpact, [violationofconfidentiality]=@violationofconfidentiality, " +
                            "[IntegrityViolation]=@IntegrityViolation, [ViolationofAccessibility]=@ViolationofAccessibility WHERE [id]=@id", connection);
                    SS.Parameters.AddWithValue("id", q.Text);
                    SS.Parameters.AddWithValue("NameofUBI", d.Text);
                    SS.Parameters.AddWithValue("Description", b.Text);
                    SS.Parameters.AddWithValue("sourceofthethreat", m.Text);
                    SS.Parameters.AddWithValue("TheObjectOfTheThreatImpact", n.Text);
                    SS.Parameters.AddWithValue("violationofconfidentiality", x.Text);
                    SS.Parameters.AddWithValue("IntegrityViolation", z.Text);
                    SS.Parameters.AddWithValue("ViolationofAccessibility", a.Text);
                    await SS.ExecuteNonQueryAsync();
                    asd.Items.Add("Успешно");
                    countupdate++;
                    qqq.Items.Clear();
                    qqq.Items.Add(countupdate);


                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                asd.Items.Add("Ошибка");

            }
        }

        private async void Button_Click_3(object sender, RoutedEventArgs e)
        {
            SqlCommand sqlCommand = new SqlCommand("SELECT * FROM [zxc] ", connection);
            await sqlCommand.ExecuteNonQueryAsync();
            SqlDataAdapter dataAdp = new SqlDataAdapter(sqlCommand);
            DataTable dt = new DataTable("SALAM");
            dataAdp.Fill(dt);
            mygrid.ItemsSource = dt.DefaultView;
        }

        private async void Button_Click_4(object sender, RoutedEventArgs e)
        {
            SqlCommand command = new SqlCommand("SELECT * FROM zxc", connection);
            SqlDataReader sqlReader = null;
            try
            {
                FileStream file = new FileStream(@"C:\Users\mvideo\OneDrive\Рабочий стол\zxc.txt", FileMode.Append); // путь к файлу 
                StreamWriter stream = new StreamWriter(file);
                try
                {
                    sqlReader = await command.ExecuteReaderAsync();
                    while (await sqlReader.ReadAsync())
                    {
                        stream.WriteLine($"({Convert.ToString(sqlReader["id"])}) ({Convert.ToString(sqlReader["NameofUBI"])}) ({Convert.ToString(sqlReader["Description"])}) " +
                                    $"({Convert.ToString(sqlReader["sourceofthethreat"])}) ({Convert.ToString(sqlReader["TheObjectOfTheThreatImpact"])}) ({Convert.ToString(sqlReader["violationofconfidentiality"])}) " +
                                    $"({Convert.ToString(sqlReader["IntegrityViolation"])}) ({Convert.ToString(sqlReader["ViolationofAccessibility"])})");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    stream.Close();
                    file.Close();
                    if (sqlReader != null)
                        sqlReader.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void Button_Click_5(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range xlRange;

            int xlRow;
            string strFileName;
            OpenFileDialog openFD = new OpenFileDialog();

            try
            {
                openFD.Filter = "Excel office |*.xls; *xlsx";
                openFD.ShowDialog();
                strFileName = openFD.FileName;

                if (strFileName != null)
                {
                    SqlCommand command;
                    xlApp = new Excel.Application();

                    try
                    {
                        command = new SqlCommand("DELETE FROM zxc WHERE [id] < 100000", connection);

                        await command.ExecuteNonQueryAsync();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    finally
                    {
                        xlWorkbook = xlApp.Workbooks.Open(strFileName);
                        xlWorksheet = xlWorkbook.Worksheets["Sheet"];
                        xlRange = xlWorksheet.UsedRange;
                        for (xlRow = 3; xlRow <= xlRange.Rows.Count; xlRow++)
                        {
                            command = new SqlCommand("INSERT INTO zxc ( NameofUBI, Description, sourceofthethreat, TheObjectOfTheThreatImpact, " +
                           "violationofconfidentiality, IntegrityViolation, ViolationofAccessibility)VALUES( @NameofUBI, @Description, @sourceofthethreat," +
                           "@TheObjectOfTheThreatImpact, @violationofconfidentiality, @IntegrityViolation, @ViolationofAccessibility)", connection);

                            command.Parameters.AddWithValue("@NameofUBI", xlRange.Cells[xlRow, 2].Text);
                            command.Parameters.AddWithValue("@Description", xlRange.Cells[xlRow, 3].Text);
                            command.Parameters.AddWithValue("@sourceofthethreat", xlRange.Cells[xlRow, 4].Text);
                            command.Parameters.AddWithValue("@TheObjectOfTheThreatImpact", xlRange.Cells[xlRow, 5].Text);
                            command.Parameters.AddWithValue("@violationofconfidentiality", xlRange.Cells[xlRow, 6].Text);
                            command.Parameters.AddWithValue("@IntegrityViolation", xlRange.Cells[xlRow, 7].Text);
                            command.Parameters.AddWithValue("@ViolationofAccessibility", xlRange.Cells[xlRow, 8].Text);

                            await command.ExecuteNonQueryAsync();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TabControl_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {

        }
    }
}
