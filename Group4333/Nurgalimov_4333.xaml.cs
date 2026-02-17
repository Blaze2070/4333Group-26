using System;
using System.Collections.Generic;
using System.IO;
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
using Npgsql;
using OfficeOpenXml;
using BCrypt.Net;

namespace Group4333
{
    public partial class Nurgalimov_4333 : Window
    {
        string connString = "Host=localhost;Port=5432;Database=3labIspro;Username=postgres;Password=2007;";

        public Nurgalimov_4333()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("4333Nurgalimov");
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "Excel файлы|*.xlsx";
            dlg.Title = "Выберите файл 5.xlsx";

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(dlg.FileName)))
                    {
                        var sheet = package.Workbook.Worksheets[0];
                        int rows = sheet.Dimension.Rows;

                        using (var conn = new NpgsqlConnection(connString))
                        {
                            conn.Open();

                            using (var clearCmd = new NpgsqlCommand("DELETE FROM Employees", conn))
                            {
                                clearCmd.ExecuteNonQuery();
                            }

                            for (int row = 2; row <= rows; row++)
                            {
                                string role = sheet.Cells[row, 1].Text?.Trim();    
                                string fullName = sheet.Cells[row, 2].Text?.Trim();
                                string login = sheet.Cells[row, 3].Text?.Trim();
                                string plainPassword = sheet.Cells[row, 4].Text?.Trim(); 

                                if (string.IsNullOrEmpty(role)) continue;

                                string password = BCrypt.Net.BCrypt.HashPassword(plainPassword);

                                using (var cmd = new NpgsqlCommand(
                                    "INSERT INTO Employees (Role, Username, Login, Password) VALUES (@role, @name, @login, @pass)",
                                    conn))
                                {
                                    cmd.Parameters.AddWithValue("@role", role);
                                    cmd.Parameters.AddWithValue("@name", fullName ?? "");
                                    cmd.Parameters.AddWithValue("@login", login ?? "");
                                    cmd.Parameters.AddWithValue("@pass", password ?? "");
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }
                    }

                    txtStatus.Text = "Импорт завершен!";
                    MessageBox.Show($"Данные успешно импортированы!", "Успех");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка импорта: " + ex.Message);
                }
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.Filter = "Excel файлы|*.xlsx";
            dlg.FileName = "Nurgalimov4333.xlsx";

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    using (var package = new ExcelPackage())
                    {
                        var data = new Dictionary<string, List<Tuple<string, string>>>();

                        using (var conn = new NpgsqlConnection(connString))
                        {
                            conn.Open();

                            var roles = new List<string>();
                            using (var rolesCmd = new NpgsqlCommand(
                                "SELECT DISTINCT Role FROM Employees WHERE Role IS NOT NULL AND Role != '' ORDER BY Role",
                                conn))
                            using (var rolesReader = rolesCmd.ExecuteReader())
                            {
                                while (rolesReader.Read())
                                {
                                    roles.Add(rolesReader[0].ToString());
                                }
                            }

                            foreach (var role in roles)
                            {
                                data[role] = new List<Tuple<string, string>>();

                                using (var cmd = new NpgsqlCommand(
                                    "SELECT Login, Password FROM Employees WHERE Role = @role ORDER BY Login",
                                    conn))
                                {
                                    cmd.Parameters.AddWithValue("@role", role);
                                    using (var reader = cmd.ExecuteReader())
                                    {
                                        while (reader.Read())
                                        {
                                            data[role].Add(Tuple.Create(
                                                reader["Login"].ToString(),
                                                reader["Password"].ToString()
                                            ));
                                        }
                                    }
                                }
                            }
                        }

                        foreach (var role in data.Keys)
                        {
                            string sheetName = role.Length > 30 ? role.Substring(0, 30) : role;
                            foreach (char c in Path.GetInvalidFileNameChars())
                            {
                                sheetName = sheetName.Replace(c, '_');
                            }

                            var worksheet = package.Workbook.Worksheets.Add(sheetName);

                            worksheet.Cells[1, 1].Value = "Логин";
                            worksheet.Cells[1, 2].Value = "Пароль";


                            int row = 2;
                            foreach (var emp in data[role])
                            {
                                worksheet.Cells[row, 1].Value = emp.Item1; 
                                worksheet.Cells[row, 2].Value = emp.Item2; 
                                row++;
                            }

                            if (row > 2)
                            {
                                worksheet.Cells[1, 1, row - 1, 2].AutoFitColumns();
                            }
                        }
                        File.WriteAllBytes(dlg.FileName, package.GetAsByteArray());
                    }

                    txtStatus.Text = $"Экспорт завершен! Файл: {dlg.FileName}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка экспорта: " + ex.Message);
                }
            }
        }
    }
}
