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
    public partial class Sales : Form
    {
        readonly SqlConnection connection = new SqlConnection(Program.dataSource);
        public Sales()
        {
            InitializeComponent();
            DataGridView.InitDefault();
            fillAgentData();
            fillGoodsData();
            LoadSalesData();
            DataGridView.ReadOnly = true;
            DataGridView.CellClick += DataGridView1_CellClick;
            ClearInputFields();
            DataGridView.Columns["SaleID"].Visible = false;
            DataGridView.Columns[1].HeaderText = "Agent adı";
            DataGridView.Columns[2].HeaderText = "Mal adı";
            DataGridView.Columns[3].HeaderText = "Say";
            DataGridView.Columns[4].HeaderText = "Cəm";
            DataGridView.Columns[5].HeaderText = "Bonus";
            DataGridView.Columns[6].HeaderText = "Qeydlər";
            DataGridView.Columns[7].HeaderText = "Tarix";
            DataGridView.Columns[8].HeaderText = "Silinmə Tarixi";
            this.KeyPreview = true;
            this.KeyDown += Agent_KeyDown;
        }

        private void ClearInputFields()
        {
            dateTimePicker1.Value = DateTime.Now;
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            textBox1.Clear();
            textBox4.Clear();
            textBox3.Clear();
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
            comboBox3.DataSource = dt;
            comboBox2.DisplayMember = "Name";
            comboBox2.ValueMember = "GoodsID";
            comboBox3.ValueMember = "Price";
        }

        private void LoadSalesData()
        {
            connection.Open();
            string query = "SELECT * FROM SalesView";
            SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
            DataTable dt = new DataTable();
            adapter.Fill(dt);
            DataGridView.DataSource = dt;
            connection.Close();
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && DataGridView.SelectedRows.Count > 0)
            {
                DataGridViewRow selectedRow = DataGridView.SelectedRows[0];
                dateTimePicker1.Value = Convert.ToDateTime(selectedRow.Cells["CreatedTime"].Value);
                comboBox1.Text = selectedRow.Cells["AgentName"].Value.ToString();
                comboBox2.Text = selectedRow.Cells["GoodsName"].Value.ToString();
                textBox1.Text = selectedRow.Cells["Quantity"].Value.ToString();
                textBox3.Text = selectedRow.Cells["Bonus"].Value.ToString();
                textBox4.Text = selectedRow.Cells["Notes"].Value.ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                if (comboBox1.SelectedIndex == -1 || comboBox2.SelectedIndex == -1)
                {
                    MessageBox.Show("Agent adı və ya malı seçməlisiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

                SqlCommand agentExistsCommand = new SqlCommand("SELECT COUNT(*) FROM Sales WHERE AgentId = @AgentId AND GoodsId = @GoodsId", connection);
                agentExistsCommand.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                agentExistsCommand.Parameters.AddWithValue("@GoodsId", comboBox2.SelectedValue);

                int agentExists = (int)agentExistsCommand.ExecuteScalar();

                //if (agentExists > 0)
                //{
                //    string updateSql = "UPDATE Sales SET Quantity = Quantity + @Quantity, Sum = Sum + @Sum, Bonus = Bonus + @Bonus, CreatedTime = @CreatedTime WHERE AgentId = @AgentId AND GoodsId = @GoodsId";

                //    using (SqlCommand updateCmd = new SqlCommand(updateSql, connection))
                //    {
                //        updateCmd.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);

                //        decimal selectedPrice;
                //        if (decimal.TryParse(textBox1.Text, out decimal quantity))
                //        {
                //            updateCmd.Parameters.AddWithValue("@Quantity", quantity);
                //            selectedPrice = Convert.ToDecimal(comboBox3.SelectedValue);
                //            updateCmd.Parameters.AddWithValue("@Sum", selectedPrice * quantity);
                //            updateCmd.Parameters.AddWithValue("@Bonus", Convert.ToDecimal(textBox3.Text) * quantity);
                //        }
                //        else
                //        {
                //            updateCmd.Parameters.AddWithValue("@Quantity", 0);
                //            updateCmd.Parameters.AddWithValue("@Sum", 0);
                //            updateCmd.Parameters.AddWithValue("@Bonus", 0);
                //        }

                //        updateCmd.Parameters.AddWithValue("@GoodsId", comboBox2.SelectedValue);
                //        updateCmd.Parameters.AddWithValue("@CreatedTime", DateTime.Now);

                //        int rowsUpdated = updateCmd.ExecuteNonQuery();

                //        if (rowsUpdated > 0)
                //        {
                //            string updateSqlForPayment = "UPDATE PaymentsModule SET Debt = Debt + @Sum1, TotalDebt = TotalDebt + @Sum1 - @Bonus1 WHERE AgentId = @AgentId AND GoodsId = @GoodsId";

                //            using (SqlCommand updateCmd1 = new SqlCommand(updateSqlForPayment, connection))
                //            {
                //                updateCmd1.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                //                updateCmd1.Parameters.AddWithValue("@GoodsId", comboBox2.SelectedValue);

                //                decimal selectedPrice1;
                //                if (decimal.TryParse(textBox1.Text, out decimal quantity1))
                //                {
                //                    selectedPrice1 = Convert.ToDecimal(comboBox3.SelectedValue);
                //                    updateCmd1.Parameters.AddWithValue("@Sum1", selectedPrice1 * quantity1);
                //                    updateCmd1.Parameters.AddWithValue("@Bonus1", Convert.ToDecimal(textBox3.Text) * quantity);
                //                }
                //                else
                //                {
                //                    updateCmd1.Parameters.AddWithValue("@Sum1", 0);
                //                    updateCmd1.Parameters.AddWithValue("@Bonus1", 0);
                //                }
                //                int rowsUpdated1 = updateCmd1.ExecuteNonQuery();
                //            }

                //            MessageBox.Show("Məlumat uğurla yeniləndi.", "Uğurlu Əməliyyat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //        }
                //        else
                //        {
                //            MessageBox.Show("Məlumat yenilənmədi.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //        }
                //    }
                //}
                //else
                //{
                    string sql = "INSERT INTO Sales (CreatedTime, AgentId, DeletedDate, Quantity, Sum, Bonus, Notes, GoodsId) VALUES " +
                        "(@CreatedTime, @AgentId, @DeletedDate, @Quantity, @Sum, @Bonus, @Notes, @GoodsId)";

                    using (SqlCommand cmd = new SqlCommand(sql, connection))
                    {
                        cmd.Parameters.AddWithValue("@CreatedTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        cmd.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                        cmd.Parameters.AddWithValue("@GoodsId", comboBox2.SelectedValue);
                        cmd.Parameters.AddWithValue("@DeletedDate", DBNull.Value);

                        decimal sum, bonus, selectedPrice;

                        selectedPrice = Convert.ToDecimal(comboBox3.SelectedValue);
                        bonus = string.IsNullOrWhiteSpace(textBox3.Text) ? 0 : Convert.ToDecimal(textBox3.Text);
                        if (decimal.TryParse(textBox1.Text, out decimal quantity))
                        {
                            cmd.Parameters.AddWithValue("@Quantity", quantity);
                            cmd.Parameters.AddWithValue("@Sum", selectedPrice * quantity);
                            cmd.Parameters.AddWithValue("@Bonus", bonus * quantity);
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue("@Quantity", 0);
                            cmd.Parameters.AddWithValue("@Sum", 0);
                            cmd.Parameters.AddWithValue("@Bonus", 0);
                        }

                        sum = selectedPrice * quantity;

                        cmd.Parameters.AddWithValue("@Notes", textBox4.Text);

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
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xəta: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                connection.Close();
                LoadSalesData();
                ClearInputFields();
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (DataGridView.SelectedRows.Count == 1)
            {
                int saleID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["SaleID"].Value);

                DateTime? deletedDate = GetSalesDeletedDate(saleID);

                if (deletedDate == null)
                {
                    if (MessageBox.Show("Bu satışı yeniləmək istədiyinizə əminsinizmi?", "Təsdiq", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
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

                            SqlCommand goodsCheckCommand = new SqlCommand("SELECT DeletedDate FROM Goods WHERE GoodsId = @GoodsId", connection);
                            goodsCheckCommand.Parameters.AddWithValue("@GoodsId", comboBox2.SelectedValue);
                            object goodsDeletedDate = goodsCheckCommand.ExecuteScalar();

                            if (goodsDeletedDate != DBNull.Value)
                            {
                                MessageBox.Show("Seçilmiş mal silinmişdir. Silinməmiş malı seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            string sql = "UPDATE Sales SET CreatedTime = @CreatedTime, AgentId = @AgentId, DeletedDate = @DeletedDate," +
                                "Quantity = @Quantity, Sum = @Sum, Bonus = @Bonus, Notes = @Notes ,GoodsID = @GoodsID WHERE SaleID = @SaleID";

                            decimal selectedPrice = Convert.ToDecimal(comboBox3.SelectedValue);
                            decimal bonus = Convert.ToDecimal(textBox3.Text);
                            using (SqlCommand cmd = new SqlCommand(sql, connection))
                            {
                                cmd.Parameters.AddWithValue("@SaleID", saleID);
                                cmd.Parameters.AddWithValue("@CreatedTime", dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss"));
                                cmd.Parameters.AddWithValue("@AgentName", comboBox1.Text);
                                cmd.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                                cmd.Parameters.AddWithValue("@GoodsName", comboBox2.Text);
                                cmd.Parameters.AddWithValue("@GoodsID", comboBox2.SelectedValue);
                                cmd.Parameters.AddWithValue("@DeletedDate", DBNull.Value);

                                if (decimal.TryParse(textBox1.Text, out decimal quantity))
                                {
                                    cmd.Parameters.AddWithValue("@Quantity", quantity);
                                    cmd.Parameters.AddWithValue("@Sum", selectedPrice * quantity);
                                    cmd.Parameters.AddWithValue("@Bonus", bonus * quantity);
                                }
                                else
                                {
                                    cmd.Parameters.AddWithValue("@Quantity", 0);
                                    cmd.Parameters.AddWithValue("@Sum", 0);
                                    cmd.Parameters.AddWithValue("@Bonus", 0);
                                }
                                cmd.Parameters.AddWithValue("@Notes", textBox4.Text);

                                int rowsAffected = cmd.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    string updateSqlForPayment = "UPDATE PaymentsModule SET Debt =  @Sum1 - Amount, TotalDebt = @Sum1 - @Bonus1 WHERE AgentId = @AgentId AND GoodsId = @GoodsId";

                                    using (SqlCommand updateCmd1 = new SqlCommand(updateSqlForPayment, connection))
                                    {
                                        updateCmd1.Parameters.AddWithValue("@AgentId", comboBox1.SelectedValue);
                                        updateCmd1.Parameters.AddWithValue("@GoodsId", comboBox2.SelectedValue);

                                        decimal selectedPrice1;
                                        if (decimal.TryParse(textBox1.Text, out decimal quantity1))
                                        {
                                            selectedPrice1 = Convert.ToDecimal(comboBox3.SelectedValue);
                                            updateCmd1.Parameters.AddWithValue("@Sum1", selectedPrice1 * quantity1);
                                            updateCmd1.Parameters.AddWithValue("@Bonus1", Convert.ToDecimal(textBox3.Text) * quantity);
                                        }
                                        else
                                        {
                                            updateCmd1.Parameters.AddWithValue("@Sum1", 0);
                                            updateCmd1.Parameters.AddWithValue("@Bonus1", 0);
                                        }

                                        int rowsUpdated1 = updateCmd1.ExecuteNonQuery();
                                    }
                                    MessageBox.Show("Satış uğurla yeniləndi.", "Uğurlu Əməliyyat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    MessageBox.Show("Satış yenilənmədi.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                            LoadSalesData();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Yeniləmə əməliyyatı ləğv edildi.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Silinən satışı redaktə edə bilməzsiniz.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Xahiş edirik, DataGridView-də bir sətir seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            ClearInputFields();
        }

        private DateTime? GetSalesDeletedDate(int salesId)
        {
            connection.Open();
            SqlCommand command = new SqlCommand("SELECT DeletedDate FROM Sales WHERE SaleID = @SaleID", connection);
            command.Parameters.AddWithValue("@SaleID", salesId);
            object result = command.ExecuteScalar();
            connection.Close();
            if (result != null && result != DBNull.Value)
            {
                return Convert.ToDateTime(result);
            }
            return null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (DataGridView.SelectedRows.Count == 1)
            {
                int saleID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["SaleID"].Value);

                if (MessageBox.Show("Bu satışı silmək istədiyinizə əminsinizmi?", "Təsdiq", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        connection.Open();

                        string sql = "UPDATE Sales SET DeletedDate = @DeletedDate WHERE SaleID = @SaleID";

                        using (SqlCommand cmd = new SqlCommand(sql, connection))
                        {
                            cmd.Parameters.AddWithValue("@SaleID", saleID);
                            cmd.Parameters.AddWithValue("@DeletedDate", DateTime.UtcNow.AddHours(+4));

                            int rowsAffected = cmd.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Satış uğurla silindi.", "Uğurlu Əməliyyat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                MessageBox.Show("Satış silinmədi.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                        LoadSalesData();
                    }
                }
                else
                {
                    MessageBox.Show("Silinmə əməliyyatı ləğv edildi.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            ClearInputFields();
        }

        private void Search_Click(object sender, EventArgs e)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Satışlar");
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
                    saveFileDialog.FileName = "Satışlar.xlsx";

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
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                LoadSalesData();
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

            string searchText = textBox2.Text.Trim();

            string query;

            if (string.IsNullOrWhiteSpace(searchText))
            {
                query = "SELECT * FROM SalesView";
            }
            else
            {
                query = "SELECT * FROM SalesView WHERE AgentName LIKE @SearchText OR GoodsName LIKE @SearchText";
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
