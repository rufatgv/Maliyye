using Maliyye.AppCode.Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace Maliyye.Forms
{
    public partial class Invoice : Form
    {
        SqlConnection connection = new SqlConnection(Program.dataSource);
        public Invoice()
        {
            InitializeComponent();
            fillAgentData();
            DataGridView.InitDefault();
            LoadInvoicesData();
            DataGridView.CellClick += DataGridView1_CellClick;
            DataGridView.ReadOnly = true;
            ClearInputFields();
            DataGridView.Columns["InvoiceID"].Visible = false;
            DataGridView.Columns[1].HeaderText = "Agent adı";
            DataGridView.Columns[2].HeaderText = "Kassa sayı";
            DataGridView.Columns[3].HeaderText = "Bonus";
            DataGridView.Columns[4].HeaderText = "Aylıq say";
            DataGridView.Columns[5].HeaderText = "Aylıq məbləğ";
            DataGridView.Columns[6].HeaderText = "Cəm";
            DataGridView.Columns[7].HeaderText = "Tarix aralığı";
            DataGridView.Columns[8].HeaderText = "Qeydlər";
            DataGridView.Columns[9].HeaderText = "Yaranma tarixi";
            DataGridView.Columns[10].HeaderText = "Silinmə tarixi";
            this.KeyPreview = true;
            this.KeyDown += Agent_KeyDown;
        }

        private void ClearInputFields()
        {
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker3.Value = DateTime.Now;
            comboBox1.SelectedIndex = -1;
            textBox1.Clear();
            textBox4.Clear();
            textBox3.Clear();
            textBox2.Clear();
            textBox6.Clear();
        }

        void fillAgentData()
        {
            SqlCommand command = new SqlCommand("SELECT * FROM Agent", connection);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "AgentId";
        }

        private void Agent_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void LoadInvoicesData()
        {
            string query = "SELECT * FROM AgentInvoiceView";
            SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            DataGridView.DataSource = dt;
        }

        private void CreateInvoiceBtn(object sender, EventArgs e)
        {

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Agent adı seçilməlidir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            connection.Open();

            DateTime startDate = dateTimePicker2.Value;
            DateTime endDate = dateTimePicker3.Value;

            string sql = "INSERT INTO Invoice (CreatedDate ,AgentId, DeletedDate, CashBoxQty, Bonus, MonthlyQty, MonthlyAmount, Sum, DateRange, Notes , SendedInv , ReceivedInv) " +
                "VALUES (@CreatedTime, @AgentId,@DeletedDate, @CashBoxQty, @Bonus, @MonthlyQty, @MonthlyAmount, @Sum, @DateRange, @Note , @SendedInv , @ReceivedInv)";

            decimal receivedInv = 0;

            SqlCommand bonusCheckCommand = new SqlCommand("SELECT Bonus FROM Sales WHERE AgentId = @AgentId", connection);
            bonusCheckCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
            object bonusResult = bonusCheckCommand.ExecuteScalar();


            using (SqlCommand cmd = new SqlCommand(sql, connection))
            {
                cmd.Parameters.AddWithValue("@CreatedTime", dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                cmd.Parameters.AddWithValue("@DeletedDate", DBNull.Value);
                cmd.Parameters.AddWithValue("@DateRange", $"{startDate.ToString("yyyy-MM-dd")} - {endDate.ToString("yyyy-MM-dd")}");

                decimal.TryParse(textBox1.Text, out decimal cashBoxQty);
                decimal.TryParse(textBox3.Text, out decimal bonus);
                decimal.TryParse(textBox4.Text, out decimal monthlyQty);
                decimal.TryParse(textBox2.Text, out decimal monthlyAmount);

                decimal sendedInv = bonus * cashBoxQty;

                if (bonusResult != DBNull.Value)
                {
                    decimal bonuss = Convert.ToDecimal(bonusResult);
                    receivedInv = bonuss - sendedInv;
                }

                cmd.Parameters.AddWithValue("@CashBoxQty", cashBoxQty);
                cmd.Parameters.AddWithValue("@Bonus", bonus);
                cmd.Parameters.AddWithValue("@MonthlyQty", monthlyQty);
                cmd.Parameters.AddWithValue("@MonthlyAmount", monthlyAmount);
                cmd.Parameters.AddWithValue("@SendedInv", sendedInv);
                cmd.Parameters.AddWithValue("@ReceivedInv", receivedInv);

                cmd.Parameters.AddWithValue("@Sum", (cashBoxQty * bonus) + (monthlyQty * monthlyAmount));

                cmd.Parameters.AddWithValue("@Note", textBox6.Text);

                int rowsAffected = cmd.ExecuteNonQuery();

                if (rowsAffected > 0)
                {
                    MessageBox.Show("Qaimə əlavə olundu");
                }
                else
                {
                    MessageBox.Show("Data insertion failed.");
                }
                connection.Close();
                LoadInvoicesData();
                ClearInputFields();
            }
        }

        private void EditInvoiceBtn(object sender, EventArgs e)
        {
            if (DataGridView.SelectedRows.Count == 1)
            {
                int invoiceID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["InvoiceID"].Value);
                DateTime startDate = dateTimePicker2.Value;
                DateTime endDate = dateTimePicker3.Value;
                DateTime? deletedDate = GetInvoiceDeletedDate(invoiceID);

                if (deletedDate == null)
                {
                    if (MessageBox.Show("Bu 'Agentə ödənişlər' sətirini redakə etməyə əminsiniz mi?", "Təsdiqləmə", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            connection.Open();

                            SqlCommand agentCheckCommand = new SqlCommand("SELECT DeletedDate FROM Agent WHERE AgentId = @AgentId", connection);
                            agentCheckCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                            object agentDeletedDate = agentCheckCommand.ExecuteScalar();

                            if (agentDeletedDate != DBNull.Value)
                            {
                                MessageBox.Show("Seçilmiş agent artıq mövcud deyil. Başqa agent seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            string sql = "UPDATE Invoice SET CreatedDate = @CreatedDate, AgentId = @AgentId, DeletedDate = @DeletedDate, CashBoxQty = @CashBoxQty , Bonus = @Bonus , MonthlyQty = @MonthlyQty, " +
                                "MonthlyAmount = @MonthlyAmount, Sum = @Sum, DateRange = @DateRange,SendedInv =  @SendedInv ,ReceivedInv = @ReceivedInv, Notes = @Notes WHERE InvoiceID = @InvoiceID";

                            decimal receivedInv = 0;

                            SqlCommand bonusCheckCommand = new SqlCommand("SELECT Bonus FROM Sales WHERE AgentId = @AgentId", connection);
                            bonusCheckCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                            object bonusResult = bonusCheckCommand.ExecuteScalar();

                            using (SqlCommand cmd = new SqlCommand(sql, connection))
                            {
                                cmd.Parameters.AddWithValue("@InvoiceID", invoiceID);
                                cmd.Parameters.AddWithValue("@CreatedDate", dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss"));
                                cmd.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                                cmd.Parameters.AddWithValue("@DeletedDate", DateTime.Now);
                                cmd.Parameters.AddWithValue("@DateRange", $"{startDate.ToString("yyyy-MM-dd")} - {endDate.ToString("yyyy-MM-dd")}");

                                decimal.TryParse(textBox1.Text, out decimal cashBoxQty);
                                decimal.TryParse(textBox3.Text, out decimal bonus);
                                decimal.TryParse(textBox4.Text, out decimal monthlyQty);
                                decimal.TryParse(textBox2.Text, out decimal monthlyAmount);

                                decimal sendedInv = bonus * cashBoxQty;

                                if (bonusResult != DBNull.Value)
                                {
                                    decimal bonuss = Convert.ToDecimal(bonusResult);
                                    receivedInv = bonuss - sendedInv;
                                }

                                cmd.Parameters.AddWithValue("@CashBoxQty", cashBoxQty);
                                cmd.Parameters.AddWithValue("@Bonus", bonus);
                                cmd.Parameters.AddWithValue("@MonthlyQty", monthlyQty);
                                cmd.Parameters.AddWithValue("@MonthlyAmount", monthlyAmount);
                                cmd.Parameters.AddWithValue("@SendedInv", sendedInv);
                                cmd.Parameters.AddWithValue("@ReceivedInv", receivedInv);

                                decimal sum = (cashBoxQty * bonus) + (monthlyQty * monthlyAmount);
                                cmd.Parameters.AddWithValue("@Sum", sum);
                                cmd.Parameters.AddWithValue("@Notes", textBox6.Text);

                                int rowsAffected = cmd.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    MessageBox.Show("Agentə ödənişlər müvəffəqiyyətlə Redaktə edildi!");
                                }
                                else
                                {
                                    MessageBox.Show("Redaktə əməliyyatında problem yarandı!");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: " + ex.Message);
                        }
                        finally
                        {
                            connection.Close();
                            LoadInvoicesData();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Redaktə əməliyyatı ləğv olundu!.");
                    }
                }
                else
                {
                    MessageBox.Show("Silinən agenti redaktə edə bilməzsiniz.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Zəhmət olmasa cədvəldən uyğun sətir seçin!.");
            }
            ClearInputFields();
        }

        private DateTime? GetInvoiceDeletedDate(int invoiceID)
        {
            connection.Open();
            SqlCommand command = new SqlCommand("SELECT DeletedDate FROM Invoice WHERE InvoiceID = @InvoiceID", connection);
            command.Parameters.AddWithValue("@InvoiceID", invoiceID);
            object result = command.ExecuteScalar();
            connection.Close();
            if (result != null && result != DBNull.Value)
            {
                return Convert.ToDateTime(result);
            }
            return null;
        }
        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && DataGridView.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = DataGridView.SelectedRows[0];
                dateTimePicker1.Value = Convert.ToDateTime(selectedRow.Cells["CreatedDate"].Value);
                comboBox1.Text = selectedRow.Cells["Name"].Value.ToString();
                textBox1.Text = selectedRow.Cells["CashBoxQty"].Value.ToString();
                textBox4.Text = selectedRow.Cells["MonthlyQty"].Value.ToString();
                textBox3.Text = selectedRow.Cells["Bonus"].Value.ToString();
                textBox2.Text = selectedRow.Cells["MonthlyAmount"].Value.ToString();
                textBox6.Text = selectedRow.Cells["Notes"].Value.ToString();
            }
        }
        private void DeleteInvoiceBtn(object sender, EventArgs e)
        {
            if (DataGridView.SelectedRows.Count == 1)
            {
                int invoiceID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["InvoiceID"].Value);

                if (MessageBox.Show("Bu 'Agentə ödənişlər' sətirini silməyə əminsiniz mi?", "Təsdiqləmə", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        connection.Open();

                        string sql = "UPDATE Invoice SET DeletedDate = @DeletedDate Where InvoiceID = @InvoiceID";

                        using (SqlCommand cmd = new SqlCommand(sql, connection))
                        {
                            cmd.Parameters.AddWithValue("@InvoiceID", invoiceID);
                            cmd.Parameters.AddWithValue("@DeletedDate", DateTime.UtcNow.AddHours(+4));
                            int rowsAffected = cmd.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Agentə ödənişlər müvəffəqiyyətlə silindi!");
                                LoadInvoicesData();
                            }
                            else
                            {
                                MessageBox.Show("Silmə əməliyyatında problem yarandı!");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                        LoadInvoicesData();
                    }
                }
                else
                {
                    MessageBox.Show("Silmə əməliyyatı ləğv olundu.");
                }
            }
            else
            {
                MessageBox.Show("Zəhmət olmasa cədvəldən uyğun sətir seçin.");
            }
            ClearInputFields();
        }

        private void Invoice_Load(object sender, EventArgs e)
        {
            LoadInvoicesData();
        }

        private void ExportToExcelBtn(object sender, EventArgs e)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Qaimələr");
                using (ExcelRange headerRange = worksheet.Cells[1, 1, 1, DataGridView.Columns.Count])
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    headerRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    headerRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    headerRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                for (int i = 0; i < DataGridView.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = DataGridView.Columns[i].HeaderText;
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                }

                for (int row = 0; row < DataGridView.Rows.Count; row++)
                {
                    for (int col = 0; col < DataGridView.Columns.Count; col++)
                    {
                        object cellValue = DataGridView.Rows[row].Cells[col].Value;
                        if (cellValue != null)
                        {
                            worksheet.Cells[row + 2, col + 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row + 2, col + 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row + 2, col + 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row + 2, col + 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells[row + 2, col + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[row + 2, col + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.AliceBlue);
                            worksheet.Cells[row + 2, col + 1].Value = cellValue.ToString();
                        }
                    }
                }


                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.FileName = "Qaimələr.xlsx";

                    int fileNumber = 1;
                    string baseFileName = Path.GetFileNameWithoutExtension(saveFileDialog.FileName);
                    string fileExtension = Path.GetExtension(saveFileDialog.FileName);
                    string folderPath = Path.GetDirectoryName(saveFileDialog.FileName);

                    while (File.Exists(Path.Combine(folderPath, $"{baseFileName}({fileNumber}){fileExtension}")))
                    {
                        fileNumber++;
                    }

                    saveFileDialog.FileName = $"{baseFileName}({fileNumber}){fileExtension}";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        excelPackage.SaveAs(excelFile);
                        MessageBox.Show("Excel faylı kimi export olundu.", "Export Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox5.Text))
            {
                LoadInvoicesData();
            }

            if (e.KeyChar == (char)13)
                button6.PerformClick();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (connection.State != ConnectionState.Open)
            {
                connection.Open();
            }

            string searchText = textBox5.Text.Trim();

            string query;

            if (string.IsNullOrWhiteSpace(searchText))
            {
                query = "SELECT * FROM AgentInvoiceView";
            }
            else
            {
                query = "SELECT * FROM AgentInvoiceView WHERE Name LIKE @SearchText";
            }

            using (SqlCommand command = new SqlCommand(query, connection))
            {
                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    command.Parameters.AddWithValue("@SearchText", "%" + searchText + "%");
                }

                SqlDataAdapter sd = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                sd.Fill(dt);

                DataGridView.DataSource = dt;
            }
            connection.Close();
        }
    }
}
