using Maliyye.AppCode.Extensions;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Maliyye.Forms
{
    public partial class PaymentsToAgent : Form
    {
        readonly SqlConnection connection = new SqlConnection(Program.dataSource);
        public PaymentsToAgent()
        {
            InitializeComponent();
            dataGridView1.InitDefault();
            fillAgentData();
            fillPaymentData();
            dataGridView1.ReadOnly = true;
            dataGridView1.CellClick += DataGridView1_CellClick;
            ClearInputFields();
            this.KeyPreview = true;
            this.KeyDown += Agent_KeyDown;
        }

        private void Agent_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void ClearInputFields()
        {
            dateTimePicker1.Value = DateTime.Now;
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            textBox1.Clear();
            textBox4.Clear();
            textBox3.Clear();
            textBox2.Clear();
            textBox5.Clear();
        }

        private void PaymentsToAgent_Load(object sender, EventArgs e)
        {
            LoadPaymentsToAgentData();
            dataGridView1.Columns["PaymentsToAgentID"].Visible = false;
            dataGridView1.Columns[1].HeaderText = "Agent adı";
            dataGridView1.Columns[2].HeaderText = "Ödənişin növü";
            dataGridView1.Columns[3].HeaderText = "Kassa sayı";
            dataGridView1.Columns[4].HeaderText = "Bonus";
            dataGridView1.Columns[5].HeaderText = "Aylıq say";
            dataGridView1.Columns[6].HeaderText = "Aylıq məbləğ";
            dataGridView1.Columns[7].HeaderText = "Cəm";
            dataGridView1.Columns[8].HeaderText = "Qeyd";
            dataGridView1.Columns[9].HeaderText = "Tarix";
            dataGridView1.Columns[10].HeaderText = "Silinmə tarixi";
        }

        private void LoadPaymentsToAgentData()
        {
            string query = "SELECT * FROM AgentPaymentsView";
            SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        void fillAgentData()
        {
            string query = "SELECT * FROM Agent";
            SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "AgentID";
        }

        void fillPaymentData()
        {

            SqlCommand command = new SqlCommand("select * from Payments", connection);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "PaymentType";
            comboBox2.ValueMember = "PaymentID";
        }

        private void CreatePaymentToAgent(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                if (comboBox1.SelectedIndex == -1 || comboBox1.SelectedIndex == -1)
                {
                    MessageBox.Show("Agent adını və ya mal adını seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                SqlCommand agentCheckCommand = new SqlCommand("SELECT DeletedDate FROM Agent WHERE AgentId = @AgentId", connection);
                agentCheckCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                object agentDeletedDate = agentCheckCommand.ExecuteScalar();

                if (agentDeletedDate != DBNull.Value)
                {
                    MessageBox.Show("Seçilmiş agent artıq mövcud deyil. Başqa agent seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                SqlCommand paymentsCheckCommand = new SqlCommand("SELECT DeletedDate FROM Payments WHERE PaymentID = @PaymentID", connection);
                paymentsCheckCommand.Parameters.AddWithValue("@PaymentID", comboBox2.SelectedValue);
                object paymentsDeletedDate = paymentsCheckCommand.ExecuteScalar();
                 
                if (paymentsDeletedDate != DBNull.Value)
                {
                    MessageBox.Show("Seçilmiş ödəniş növü silinmişdir. Silinməmiş malı seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                string sql = "INSERT INTO PaymentsToAgent (CreatedDate,CashBoxQty, Bonus, MonthlyQty, MonthlyAmount,Sum, Note,DeletedDate,PaymentID,AgentID) VALUES " +
                    "(@CreatedTime, @CashBoxQty, @Bonus, @MonthlyQty, @MonthlyAmount, @Sum,@Note, @DeletedDate, @PaymentID ,  @AgentID)";
                SqlCommand cmd = new SqlCommand(sql, connection);
                cmd.Parameters.AddWithValue("@CreatedTime", dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss"));
                cmd.Parameters.AddWithValue("@PaymentMethodName", comboBox2.SelectedValue.ToString());
                cmd.Parameters.AddWithValue("@DeletedDate", DBNull.Value);
                cmd.Parameters.AddWithValue("@AgentID", comboBox1.SelectedValue);
                cmd.Parameters.AddWithValue("@PaymentID", comboBox2.SelectedValue);

                decimal.TryParse(textBox1.Text, out decimal cashBoxQty);
                decimal.TryParse(textBox3.Text, out decimal bonus);
                decimal.TryParse(textBox2.Text, out decimal monthlyQty);
                decimal.TryParse(textBox5.Text, out decimal monthlyAmount);

                cmd.Parameters.AddWithValue("@CashBoxQty", cashBoxQty);
                cmd.Parameters.AddWithValue("@Bonus", bonus);
                cmd.Parameters.AddWithValue("@MonthlyQty", monthlyQty);
                cmd.Parameters.AddWithValue("@MonthlyAmount", monthlyAmount);
                cmd.Parameters.AddWithValue("@Sum", (cashBoxQty * bonus) + (monthlyQty * monthlyAmount));
                cmd.Parameters.AddWithValue("@Note", textBox4.Text);

                int rowsAffected = cmd.ExecuteNonQuery();

                MessageBox.Show(rowsAffected > 0 ? "Əlavə olundu." : "Data insertion failed.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                connection.Close();
                LoadPaymentsToAgentData();
            }
            ClearInputFields();

        }

        private void EditPaymentToAgent(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Zəhmət olmasa cədvəldən uyğun sətir seçin!.");
                return;
            }

            int paymentsToAgentId = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["PaymentsToAgentId"].Value);
            DateTime? deletedDate = GetPaymentsToAgentDeletedDate(paymentsToAgentId);
            if (deletedDate == null)
            {
                if (MessageBox.Show("Bu 'Agentə ödənişlər' sətirini redakə etməyə əminsiniz mi?", "Təsdiqləmə", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    MessageBox.Show("Redaktə əməliyyatı ləğv olundu!.");
                    return;
                }

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

                    SqlCommand paymentsCheckCommand = new SqlCommand("SELECT DeletedDate FROM Payments WHERE PaymentID = @PaymentID", connection);
                    paymentsCheckCommand.Parameters.AddWithValue("@PaymentID", comboBox2.SelectedValue);
                    object paymentsDeletedDate = paymentsCheckCommand.ExecuteScalar();

                    if (paymentsDeletedDate != DBNull.Value)
                    {
                        MessageBox.Show("Seçilmiş ödəniş növü silinmişdir. Silinməmiş malı seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }


                    string sql = "UPDATE PaymentsToAgent SET CreatedDate = @CreatedDate,CashBoxQty = @CashBoxQty , Bonus = @Bonus , MonthlyQty = @MonthlyQty, " +
                        "MonthlyAmount = @MonthlyAmount, Sum = @Sum, Note = @Note , DeletedDate = @DeletedDate,PaymentID = @PaymentID , AgentID = @AgentID WHERE PaymentsToAgentId = @PaymentsToAgentId";
                    SqlCommand cmd = new SqlCommand(sql, connection);
                    cmd.Parameters.AddWithValue("@PaymentsToAgentId", paymentsToAgentId);
                    cmd.Parameters.AddWithValue("@CreatedDate", dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss"));
                    cmd.Parameters.AddWithValue("@AgentID", comboBox1.SelectedValue);
                    cmd.Parameters.AddWithValue("@PaymentID", comboBox2.SelectedValue);
                    cmd.Parameters.AddWithValue("@DeletedDate", DBNull.Value);

                    decimal.TryParse(textBox1.Text, out decimal cashBoxQty);
                    decimal.TryParse(textBox3.Text, out decimal bonus);
                    decimal.TryParse(textBox2.Text, out decimal monthlyQty);
                    decimal.TryParse(textBox5.Text, out decimal monthlyAmount);

                    cmd.Parameters.AddWithValue("@CashBoxQty", cashBoxQty);
                    cmd.Parameters.AddWithValue("@Bonus", bonus);
                    cmd.Parameters.AddWithValue("@MonthlyQty", monthlyQty);
                    cmd.Parameters.AddWithValue("@MonthlyAmount", monthlyAmount);
                    cmd.Parameters.AddWithValue("@Sum", (cashBoxQty * bonus) + (monthlyQty * monthlyAmount));
                    cmd.Parameters.AddWithValue("@Note", textBox4.Text);

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
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    connection.Close();
                    LoadPaymentsToAgentData();
                }
                ClearInputFields();
            }
            else
            {
                MessageBox.Show("Silinən agenti redaktə edə bilməzsiniz.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private DateTime? GetPaymentsToAgentDeletedDate(int paymentsToAgentId)
        {
            connection.Open();
            SqlCommand command = new SqlCommand("SELECT DeletedDate FROM PaymentsToAgent WHERE PaymentsToAgentID = @PaymentsToAgentID", connection);
            command.Parameters.AddWithValue("@PaymentsToAgentID", paymentsToAgentId);
            object result = command.ExecuteScalar();
            connection.Close();
            if (result != null && result != DBNull.Value)
            {
                return Convert.ToDateTime(result);
            }
            return null;
        }


        private void DeletePaymentToAgent(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 1)
            {
                MessageBox.Show("Zəhmət olmasa cədvəldən uyğun sətir seçin!.");
                return;
            }

            int paymentsToAgentId = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["PaymentsToAgentId"].Value);

            if (MessageBox.Show("Bu 'Agentə ödənişlər' sətirini silməyə əminsiniz mi?", "Təsdiqləmə", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                MessageBox.Show("Silinmə əməliyyatı ləğv olundu.");
                return;
            }

            try
            {
                connection.Open();

                SqlCommand command = new SqlCommand("Update PaymentsToAgent SET DeletedDate = @DeletedDate WHERE PaymentsToAgentId = @PaymentsToAgentId", connection);
                command.Parameters.AddWithValue("@PaymentsToAgentId", paymentsToAgentId);
                command.Parameters.AddWithValue("@DeletedDate", DateTime.UtcNow.AddHours(+4));

                int rowsAffected = command.ExecuteNonQuery();

                if (rowsAffected > 0)
                {
                    MessageBox.Show("'Agentə ödənişlər' müvəffəqiyyətlə silindi!.");
                }
                else
                {
                    MessageBox.Show("'Agentə ödənişlər' silinmə əməliyyatında problem yarandı!");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                connection.Close();
                LoadPaymentsToAgentData();
            }
            ClearInputFields();
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && dataGridView1.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];
                dateTimePicker1.Value = Convert.ToDateTime(selectedRow.Cells["CreatedDate"].Value);
                comboBox1.Text = selectedRow.Cells["AgentName"].Value.ToString();
                comboBox2.Text = selectedRow.Cells["PaymentType"].Value.ToString();
                textBox1.Text = selectedRow.Cells["CashBoxQty"].Value.ToString();
                textBox3.Text = selectedRow.Cells["Bonus"].Value.ToString();
                textBox2.Text = selectedRow.Cells["MonthlyQty"].Value.ToString();
                textBox5.Text = selectedRow.Cells["MonthlyAmount"].Value.ToString();
                textBox4.Text = selectedRow.Cells["Note"].Value.ToString();
            }
        }

        private void ExportToExcelBtn(object sender, EventArgs e)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Agentə Ödənişlər");
                using (ExcelRange headerRange = worksheet.Cells[1, 1, 1, dataGridView1.Columns.Count])
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    headerRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    headerRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    headerRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    headerRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                }

                for (int row = 0; row < dataGridView1.Rows.Count; row++)
                {
                    for (int col = 0; col < dataGridView1.Columns.Count; col++)
                    {
                        object cellValue = dataGridView1.Rows[row].Cells[col].Value;
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
                    saveFileDialog.FileName = "Agentə-ödəniş.xlsx";

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

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox6.Text))
            {
                LoadPaymentsToAgentData();
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

            string searchText = textBox6.Text.Trim();

            string query;

            if (string.IsNullOrWhiteSpace(searchText))
            {
                query = "SELECT * FROM AgentPaymentsView";
            }
            else
            {
                query = "SELECT * FROM AgentPaymentsView WHERE AgentName LIKE @SearchText";
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

                dataGridView1.DataSource = dt;
            }
            connection.Close();
        }
    }
}
