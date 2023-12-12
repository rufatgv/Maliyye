using Maliyye.AppCode.Extensions;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;

namespace Maliyye.Forms
{
    public partial class PaymentsModule : Form
    {
        readonly SqlConnection connection = new SqlConnection(Program.dataSource);
        public PaymentsModule()
        {
            InitializeComponent();
            fillAgentData();
            fillGoodsData();
            fillPaymentTypeData();
            DataGridView.InitDefault();
            LoadPaymentsModuleData();
            DataGridView.CellClick += DataGridView_CellClick;
            DataGridView.ReadOnly = true;
            ClearInputFields();
            DataGridView.Columns[1].HeaderText = "Agent adı";
            DataGridView.Columns[2].HeaderText = "Mal adı";
            DataGridView.Columns[3].HeaderText = "Ödənişin növü";
            DataGridView.Columns[4].HeaderText = "Məbləğ";
            DataGridView.Columns[5].HeaderText = "Qeyd";
            DataGridView.Columns[8].HeaderText = "Tarix";
            DataGridView.Columns[9].HeaderText = "Silinmə tarixi";
            DataGridView.Columns["Debt"].Visible = false;
            DataGridView.Columns["TotalDebt"].Visible = false;
            DataGridView.Columns["Korreksiya"].Visible = false;
            DataGridView.Columns["PaymentModuleID"].Visible = false;
            this.KeyPreview = true;
            this.KeyDown += Agent_KeyDown;

        }

        private void ClearInputFields()
        {
            dateTimePicker1.Value = DateTime.Now;
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            textBox1.Clear();
            textBox2.Clear();
        }
        private void Agent_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
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

        void fillGoodsData()
        {
            SqlCommand command = new SqlCommand("select * from Goods", connection);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "GoodsID";
        }

        void fillPaymentTypeData()
        {
            SqlCommand command = new SqlCommand("select * from Payments", connection);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            comboBox3.DataSource = dt;
            comboBox3.DisplayMember = "PaymentType";
            comboBox3.ValueMember = "PaymentID";
        }

        private void LoadPaymentsModuleData()
        {
            connection.Open();
            string query = "SELECT * FROM PaymentDetailsView";
            SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            DataGridView.DataSource = dt;
            connection.Close();
        }

        private void CreatePaymentModuleBtn(object sender, EventArgs e)
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

                SqlCommand goodsCheckCommand = new SqlCommand("SELECT DeletedDate FROM Goods WHERE GoodsID = @GoodsID", connection);
                goodsCheckCommand.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                object goodsDeletedDate = goodsCheckCommand.ExecuteScalar();

                if (goodsDeletedDate != DBNull.Value)
                {
                    MessageBox.Show("Seçilmiş mal silinmişdir. Silinməmiş malı seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                SqlCommand paymentsCheckCommand = new SqlCommand("SELECT DeletedDate FROM Payments WHERE PaymentID = @PaymentID", connection);
                paymentsCheckCommand.Parameters.AddWithValue("@PaymentID", comboBox3.SelectedValue);
                object paymentsDeletedDate = paymentsCheckCommand.ExecuteScalar();

                if (paymentsDeletedDate != DBNull.Value)
                {
                    MessageBox.Show("Seçilmiş ödəniş növü silinmişdir. Silinməmiş malı seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                SqlCommand paymentExistsCommand = new SqlCommand("SELECT COUNT(*) FROM PaymentsModule WHERE AgentID = @AgentID AND GoodsID = @GoodsID AND PaymentID=@PaymentID", connection);
                paymentExistsCommand.Parameters.AddWithValue("@AgentID", comboBox1.SelectedValue);
                paymentExistsCommand.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                paymentExistsCommand.Parameters.AddWithValue("@PaymentID", comboBox3.SelectedValue);

                int paymentExists = (int)paymentExistsCommand.ExecuteScalar();

                string sql = "INSERT INTO PaymentsModule (CreatedTime,GoodsID, PaymentID, Amount, Notes ,AgentID, DeletedDate,Debt , TotalDebt) VALUES " +
                    "(@CreatedTime,@GoodsID, @PaymentID, @Amount, @Notes, @AgentID, @DeletedDate, @Debt , @TotalDebt )";

                decimal debt = 0;
                decimal totalDebt = 0;

                SqlCommand sumCheckCommand = new SqlCommand("SELECT SUM(Sum) FROM Sales WHERE AgentId = @AgentId AND GoodsID = @GoodsID AND DeletedDate IS NUll", connection);
                sumCheckCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                sumCheckCommand.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                object sumResult = sumCheckCommand.ExecuteScalar();

                string query = "SELECT Debt " +
                  "FROM PaymentsModule " +
                  "WHERE AgentID = @AgentID AND GoodsID = @GoodsID " +
                  "ORDER BY PaymentModuleID DESC ";

                using (SqlDataAdapter adapter = new SqlDataAdapter(query, connection))
                {
                    adapter.SelectCommand.Parameters.AddWithValue("@AgentID", comboBox1.SelectedValue);
                    adapter.SelectCommand.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);

                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        debt = Convert.ToDecimal(dt.Rows[0]["Debt"]);
                    }
                }


                SqlCommand bonusCheckCommand = new SqlCommand("SELECT SUM(Bonus) FROM Sales WHERE AgentId = @AgentId AND GoodsID = @GoodsID AND DeletedDate is Null", connection);
                bonusCheckCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                bonusCheckCommand.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                object bonusResult = bonusCheckCommand.ExecuteScalar();

                if (sumResult != DBNull.Value)
                {
                    decimal sum = Convert.ToDecimal(sumResult);

                    if (decimal.TryParse(textBox1.Text, out decimal amount))
                    {
                        if (debt > 0)
                        {
                            debt -= amount;
                        }
                        else
                        {
                            debt = sum - amount;
                        }
                    }
                }
                if (bonusResult != DBNull.Value)
                {
                    decimal bonus = Convert.ToDecimal(bonusResult);
                    totalDebt = debt - bonus;
                }

                using (SqlCommand cmd = new SqlCommand(sql, connection))
                {
                    cmd.Parameters.AddWithValue("@CreatedTime", dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss"));
                    cmd.Parameters.AddWithValue("@AgentID", comboBox1.SelectedValue);
                    cmd.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                    cmd.Parameters.AddWithValue("@PaymentID", comboBox3.SelectedValue);
                    cmd.Parameters.AddWithValue("@DeletedDate", DBNull.Value);
                    cmd.Parameters.AddWithValue("@Debt", debt);
                    cmd.Parameters.AddWithValue("@TotalDebt", totalDebt);

                    if (decimal.TryParse(textBox1.Text, out decimal amount))
                    {
                        cmd.Parameters.AddWithValue("@Amount", amount);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@Amount", 0);
                    }

                    cmd.Parameters.AddWithValue("@Notes", textBox2.Text);

                    int rowsAffected = cmd.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Məlumat uğurla əlavə edildi.", "Uğurlu Əməliyyat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Məlumat əlavə edilmədi.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                connection.Close();
                LoadPaymentsModuleData();
            }
        }

        private void EditPaymentModuleBtn(object sender, EventArgs e)
        {
            int PaymentModuleID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["PaymentModuleID"].Value);
            DateTime? deletedDate = GetPaymentModuleDeletedDate(PaymentModuleID);

            if (deletedDate == null)
            {
                if (MessageBox.Show("Bu ödənişi yeniləmək istədiyinizdən əminsiniz?", "Təsdiq", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    MessageBox.Show("Yeniləmə əməliyyatı ləğv edildi.", "Ləğv", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                try
                {
                    connection.Open();

                    if (comboBox1.SelectedIndex == -1 || comboBox2.SelectedIndex == -1 || comboBox3.SelectedIndex == -1)
                    {
                        MessageBox.Show("Agent adı, mal adı və ödənişin növü seçilməlidir.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

                    SqlCommand goodsCheckCommand = new SqlCommand("SELECT DeletedDate FROM Goods WHERE GoodsID = @GoodsID", connection);
                    goodsCheckCommand.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                    object goodsDeletedDate = goodsCheckCommand.ExecuteScalar();

                    if (goodsDeletedDate != DBNull.Value)
                    {
                        MessageBox.Show("Seçilmiş mal silinmişdir. Silinməmiş malı seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    SqlCommand paymentsCheckCommand = new SqlCommand("SELECT DeletedDate FROM Payments WHERE PaymentID = @PaymentID", connection);
                    paymentsCheckCommand.Parameters.AddWithValue("@PaymentID", comboBox3.SelectedValue);
                    object paymentsDeletedDate = paymentsCheckCommand.ExecuteScalar();

                    if (paymentsDeletedDate != DBNull.Value)
                    {
                        MessageBox.Show("Seçilmiş ödəniş növü silinmişdir. Silinməmiş malı seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    string sql = "UPDATE PaymentsModule SET CreatedTime = @CreatedTime, GoodsID = @GoodsID , PaymentID = @PaymentID," +
                        "Amount = @Amount, Notes = @Notes, AgentId = @AgentId , DeletedDate = @DeletedDate,  Debt = @Debt ,TotalDebt = @TotalDebt  WHERE PaymentModuleID = @PaymentModuleID";

                    decimal debt = 0;
                    decimal totalDebt = 0;
                    decimal debt1 = 0;

                    SqlCommand sumCheckCommand = new SqlCommand("SELECT Sum FROM Sales WHERE AgentId = @AgentId AND GoodsID = @GoodsID", connection);
                    sumCheckCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                    sumCheckCommand.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                    object sumResult = sumCheckCommand.ExecuteScalar();

                    decimal previousAmount = 0;
                    decimal debtamountdifference = 0;

                    string query = "SELECT * FROM PaymentsModule WHERE PaymentModuleID = @PaymentModuleID";
                    using (SqlDataAdapter adapter = new SqlDataAdapter(query, connection))
                    {
                        adapter.SelectCommand.Parameters.AddWithValue("@PaymentModuleID", PaymentModuleID);

                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            debt = Convert.ToDecimal(dt.Rows[0]["Debt"].ToString());

                            previousAmount = Convert.ToDecimal(dt.Rows[0]["Amount"].ToString());

                            debt1 = previousAmount - Convert.ToDecimal(textBox1.Text) + debt;

                            debtamountdifference = Convert.ToDecimal(textBox1.Text) - previousAmount;
                        }
                    }

                    decimal debtLast = 0;
                    decimal amountlast = 0;

                    string updatesql = "Select * From PaymentsModule " +
                           "WHERE AgentId=@AgentId and GoodsID=@GoodsID and  PaymentModuleID >= @PaymentModuleID";

                    using (SqlDataAdapter adapter = new SqlDataAdapter(updatesql, connection))
                    {
                        adapter.SelectCommand.Parameters.AddWithValue("@PaymentModuleID", PaymentModuleID);
                        adapter.SelectCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                        adapter.SelectCommand.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);

                        DataTable dt = new DataTable();
                        adapter.Fill(dt);

                        foreach (DataRow row in dt.Rows)
                        {
                            debtLast = Convert.ToDecimal(row["Debt"].ToString());
                            amountlast = Convert.ToDecimal(row["Amount"].ToString());

                            string updateDebtLastSql = "UPDATE PaymentsModule SET Debt = @DebtLast WHERE PaymentModuleID >= @PaymentModuleID AND AgentId = @AgentId";
                            using (SqlCommand updateDebtLastCmd = new SqlCommand(updateDebtLastSql, connection))
                            {
                                updateDebtLastCmd.Parameters.AddWithValue("@PaymentModuleID", Convert.ToInt32(row["PaymentModuleID"]));
                                updateDebtLastCmd.Parameters.AddWithValue("@DebtLast", debtLast - debtamountdifference);
                                updateDebtLastCmd.Parameters.AddWithValue("@AgentId", Convert.ToInt32(row["AgentId"]));
                                updateDebtLastCmd.Parameters.AddWithValue("@GoodsID", Convert.ToInt32(row["GoodsID"]));

                                int rowsUpdated = updateDebtLastCmd.ExecuteNonQuery();
                            }
                        }
                    }

                    SqlCommand bonusCheckCommand = new SqlCommand("SELECT Bonus FROM Sales WHERE AgentId = @AgentId AND GoodsID = @GoodsID", connection);
                    bonusCheckCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                    bonusCheckCommand.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                    object bonusResult = bonusCheckCommand.ExecuteScalar();

                    if (bonusResult != DBNull.Value)
                    {
                        decimal bonus = Convert.ToDecimal(bonusResult);
                        totalDebt = debt1 - bonus;
                    }

                    using (SqlCommand cmd = new SqlCommand(sql, connection))
                    {
                        cmd.Parameters.AddWithValue("@PaymentModuleID", PaymentModuleID);
                        cmd.Parameters.AddWithValue("@CreatedTime", dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss"));
                        cmd.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                        cmd.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                        cmd.Parameters.AddWithValue("@PaymentID", comboBox3.SelectedValue);
                        cmd.Parameters.AddWithValue("@DeletedDate", DBNull.Value);
                        cmd.Parameters.AddWithValue("@Debt", debt1);
                        cmd.Parameters.AddWithValue("@TotalDebt", totalDebt);

                        decimal amount = decimal.TryParse(textBox1.Text, out amount) ? amount : 0;
                        cmd.Parameters.AddWithValue("@Amount", amount);
                        cmd.Parameters.AddWithValue("@Notes", textBox2.Text);

                        int rowsAffected = cmd.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Ödəniş " + (rowsAffected > 0 ? "uğurla yeniləndi." : "yenilənmədi."), "Uğurlu", MessageBoxButtons.OK, rowsAffected > 0 ? MessageBoxIcon.Information : MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                    MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    connection.Close();
                    LoadPaymentsModuleData();
                }
            }
            else
            {
                MessageBox.Show("Silinən agenti redaktə edə bilməzsiniz.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        private DateTime? GetPaymentModuleDeletedDate(int PaymentModuleID)
        {
            connection.Open();
            SqlCommand command = new SqlCommand("SELECT DeletedDate FROM PaymentsModule WHERE PaymentModuleID = @PaymentModuleID", connection);
            command.Parameters.AddWithValue("@PaymentModuleID", PaymentModuleID);
            object result = command.ExecuteScalar();
            connection.Close();
            if (result != null && result != DBNull.Value)
            {
                return Convert.ToDateTime(result);
            }
            return null;
        }

        private void DeletePaymentModuleBtn(object sender, EventArgs e)
        {
            if (DataGridView.SelectedRows.Count == 1)
            {
                int paymentID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["PaymentModuleID"].Value);

                if (MessageBox.Show("Bu ödənişi silmək istədiyinizdən əminsiniz?", "Təsdiq", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        connection.Open();

                        string sql = "UPDATE PaymentsModule SET DeletedDate = @DeletedDate WHERE PaymentModuleID = @PaymentModuleID";

                        using (SqlCommand cmd = new SqlCommand(sql, connection))
                        {
                            cmd.Parameters.AddWithValue("@DeletedDate", DateTime.UtcNow.AddHours(+4));
                            cmd.Parameters.AddWithValue("@PaymentModuleID", paymentID);

                            int rowsAffected = cmd.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Ödəniş uğurla silindi.", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Ödəniş silinmədi.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Xəta: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        connection.Close();
                        LoadPaymentsModuleData();
                    }
                }
                else
                {
                    MessageBox.Show("Silinmə əməliyyatı ləğv edildi.", "Ləğv", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        private void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow selectedRow = DataGridView.Rows[e.RowIndex];
                dateTimePicker1.Value = Convert.ToDateTime(selectedRow.Cells["CreatedTime"].Value);
                comboBox1.Text = selectedRow.Cells["AgentName"].Value.ToString();
                comboBox2.Text = selectedRow.Cells["GoodsName"].Value.ToString();
                comboBox3.Text = selectedRow.Cells["PaymentType"].Value.ToString();
                textBox1.Text = selectedRow.Cells["Amount"].Value.ToString();
                textBox2.Text = selectedRow.Cells["Notes"].Value.ToString();
            }
        }

        private void Excell_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Ödənişlər");
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
                    saveFileDialog.FileName = "Ödənişlər.xlsx";

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

        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox7.Text))
            {
                LoadPaymentsModuleData();
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

            string searchText = textBox7.Text.Trim();

            string query;

            if (string.IsNullOrWhiteSpace(searchText))
            {
                query = "SELECT * FROM PaymentDetailsView";
            }
            else
            {
                query = "SELECT * FROM PaymentDetailsView WHERE AgentName LIKE @SearchText OR GoodsName LIKE @SearchText OR PaymentType LIKE @SearchText";
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
