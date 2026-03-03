using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Windows;
using BCrypt.Net;
using Npgsql;
using OfficeOpenXml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Group4333
{

    public class EmployeeData
    {
        public string Role { get; set; }
        public string FIO { get; set; }
        public string Login { get; set; }
        public string Password { get; set; }
    }

    public partial class Nurgalimov_4333 : Window
    {
        string connString = "Host=localhost;Port=5432;Database=3labIspro;Username=postgres;Password=2007;";

        public Nurgalimov_4333()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("4333Nurgalimov");
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
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

        private void BtnImportJson_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "JSON файлы|*.json";
            dlg.Title = "Выберите файл 5.json";

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    string json = File.ReadAllText(dlg.FileName);
                    var employees = JsonSerializer.Deserialize<List<EmployeeData>>(json);

                    using (var conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();

                        using (var clearCmd = new NpgsqlCommand("DELETE FROM Employees", conn))
                        {
                            clearCmd.ExecuteNonQuery();
                        }

                        foreach (var emp in employees)
                        {
                            string hashedPassword = BCrypt.Net.BCrypt.HashPassword(emp.Password);

                            using (var cmd = new NpgsqlCommand(
                                "INSERT INTO Employees (Role, Username, Login, Password) VALUES (@role, @name, @login, @pass)",
                                conn))
                            {
                                cmd.Parameters.AddWithValue("@role", emp.Role);
                                cmd.Parameters.AddWithValue("@name", emp.FIO);
                                cmd.Parameters.AddWithValue("@login", emp.Login);
                                cmd.Parameters.AddWithValue("@pass", hashedPassword);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    txtStatus.Text = "Импорт завершен!";
                    MessageBox.Show($"Импорт из JSON завершен! Загружено записей: {employees.Count}", "Успех");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка импорта JSON: " + ex.Message);
                }
            }
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
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

                            worksheet.Cells[1, 1].Value = "Login";
                            worksheet.Cells[1, 2].Value = "Password";


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

        private void BtnExportWord_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.Filter = "Word файлы|*.docx";
            dlg.FileName = "Nurgalimov4333.docx";

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    List<EmployeeData> employees = new List<EmployeeData>();

                    using (var conn = new NpgsqlConnection(connString))
                    {
                        conn.Open();
                        using (var cmd = new NpgsqlCommand(
                            "SELECT Role, Login, Password FROM Employees ORDER BY Role", conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                employees.Add(new EmployeeData
                                {
                                    Role = reader["Role"].ToString(),
                                    Login = reader["Login"].ToString(),
                                    Password = reader["Password"].ToString()
                                });
                            }
                        }
                    }

                    if (employees.Count == 0)
                    {
                        MessageBox.Show("Нет данных для экспорта");
                        return;
                    }

                    var groups = employees.GroupBy(e => e.Role).OrderBy(g => g.Key);

                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(dlg.FileName, WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                        mainPart.Document = new Document();
                        Body body = mainPart.Document.AppendChild(new Body());

                        body.AppendChild(new Paragraph());

                        int groupCount = 0;
                        foreach (var group in groups)
                        {
                            if (groupCount > 0)
                            {
                                Paragraph pageBreak = body.AppendChild(new Paragraph());
                                Run pageBreakRun = pageBreak.AppendChild(new Run());
                                pageBreakRun.AppendChild(new Break() { Type = BreakValues.Page });
                            }

                            Paragraph roleParagraph = body.AppendChild(new Paragraph());
                            roleParagraph.ParagraphProperties = new ParagraphProperties();
                            roleParagraph.ParagraphProperties.AppendChild(new Justification() { Val = JustificationValues.Center });

                            Run roleRun = roleParagraph.AppendChild(new Run());
                            roleRun.AppendChild(new Text($"Роль: {group.Key}"));
                            roleRun.RunProperties = new RunProperties();
                            roleRun.RunProperties.AppendChild(new Bold());
                            roleRun.RunProperties.AppendChild(new Underline());

                            body.AppendChild(new Paragraph());

                            Table table = new Table();

                            TableProperties tblProps = new TableProperties();
                            TableBorders tblBorders = new TableBorders();
                            tblBorders.TopBorder = new TopBorder() { Val = BorderValues.Single, Size = 4 };
                            tblBorders.BottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 4 };
                            tblBorders.LeftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 4 };
                            tblBorders.RightBorder = new RightBorder() { Val = BorderValues.Single, Size = 4 };
                            tblBorders.InsideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 2 };
                            tblBorders.InsideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 2 };
                            tblProps.AppendChild(tblBorders);
                            table.AppendChild(tblProps);

                            TableRow headerRow = new TableRow();

                            AddCell(headerRow, "Логин", true);
                            AddCell(headerRow, "Пароль", true);

                            table.AppendChild(headerRow);

                            foreach (var emp in group)
                            {
                                TableRow dataRow = new TableRow();

                                AddCell(dataRow, emp.Login, false);
                                AddCell(dataRow, emp.Password, false);

                                table.AppendChild(dataRow);
                            }

                            body.AppendChild(table);

                            Paragraph countParagraph = body.AppendChild(new Paragraph());
                            countParagraph.ParagraphProperties = new ParagraphProperties();
                            countParagraph.ParagraphProperties.AppendChild(new Justification() { Val = JustificationValues.Right });

                            body.AppendChild(new Paragraph());
                            groupCount++;
                        }

                        mainPart.Document.Save();
                    }

                    txtStatus.Text = $"Экспорт в Word завершен! Файл: {dlg.FileName}";
                    MessageBox.Show($"Экспорт в Word завершен! Файл сохранен: {dlg.FileName}", "Успех");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка экспорта в Word: " + ex.Message);
                }
            }
        }
        private void AddCell(TableRow row, string text, bool isHeader)
        {
            TableCell cell = new TableCell();
            cell.AppendChild(new Paragraph(new Run(new Text(text))));

            if (isHeader)
            {
                cell.TableCellProperties = new TableCellProperties();
                cell.TableCellProperties.AppendChild(new Shading()
                {
                    Fill = "D3D3D3",
                    Val = ShadingPatternValues.Clear
                });

                RunProperties runProps = new RunProperties();
                runProps.AppendChild(new Bold());
                cell.Descendants<Run>().First().RunProperties = runProps;
            }

            row.AppendChild(cell);
        }
    }
}
