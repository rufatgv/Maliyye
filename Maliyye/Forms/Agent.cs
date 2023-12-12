using Maliyye.AppCode.Extensions;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;


namespace Maliyye.Forms
{

    public partial class Agent : Form
    {
        SqlConnection  connection = new SqlConnection(Program.dataSource);
        public Agent()
        {
            InitializeComponent();
            DataGridView.InitDefault();
            DataGridView.CellClick += DataGridView_CellClick;
            DataGridView.ReadOnly = true;
            ClearInputFields();
            this.KeyPreview = true; 
            this.KeyDown += Agent_KeyDown;
        }
        private void Agent_Load(object sender, EventArgs e)
        {
            BindData();
            DataGridView.Columns["AgentID"].Visible = false;
            DataGridView.Columns[1].HeaderText = "Ad";
            DataGridView.Columns[2].HeaderText = "Silinmə tarixi";
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
            textBox3.Clear(); 
            textBox2.Clear(); 
        }

        void BindData()
        {
            SqlCommand command = new SqlCommand("SELECT * FROM Agent", connection);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            DataGridView.DataSource = dt;
        }

        private void CreateAgentBtn(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Agent adı daxil edin", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            connection.Open();
            SqlCommand command = new SqlCommand("INSERT INTO Agent (Name, DeletedDate) VALUES (@Name, @DeletedDate)", connection);
            command.Parameters.AddWithValue("@Name", textBox2.Text.ToString());
            command.Parameters.AddWithValue("@DeletedDate", DBNull.Value);

            command.ExecuteNonQuery();
            MessageBox.Show("Əlavə edildi", "Uğurlu Əməliyyat", MessageBoxButtons.OK, MessageBoxIcon.Information);
            connection.Close();
            BindData();
            ClearInputFields();
        }

        private void UpdateAgentBtn(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Boş ola bilməz", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (DataGridView.SelectedRows.Count == 1)
            {
                int agentID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["AgentID"].Value);
                DateTime? deletedDate = GetAgentDeletedDate(agentID);

                if (deletedDate == null)
                {
                    if (MessageBox.Show("Redaktə etmək istədiyinizə əminsiniz?", "Təsdiq", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        try
                        {
                            connection.Open();

                            SqlCommand command = new SqlCommand("UPDATE Agent SET Name = @Value, DeletedDate = @DeletedDate WHERE AgentId = @ID", connection);
                            command.Parameters.AddWithValue("@Value", textBox2.Text);
                            command.Parameters.AddWithValue("@ID", agentID);
                            command.Parameters.AddWithValue("@DeletedDate", DBNull.Value);

                            int rowsAffected = command.ExecuteNonQuery();
                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Redaktə edildi", "Uğurlu Əməliyyat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                BindData();
                            }
                            else
                            {
                                MessageBox.Show("Redaktə edilmədi", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            connection.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Redaktə əməliyyatı ləğv edildi.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Silinən agenti redaktə edə bilməzsiniz.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            BindData();
            ClearInputFields();
        }

        private DateTime? GetAgentDeletedDate(int agentID)
        {
            connection.Open();
            SqlCommand command = new SqlCommand("SELECT DeletedDate FROM Agent WHERE AgentId = @AgentId", connection);
            command.Parameters.AddWithValue("@AgentId", agentID);
            object result = command.ExecuteScalar();
            connection.Close();
            if (result != null && result != DBNull.Value)
            {
                return Convert.ToDateTime(result);
            }
            return null;
        }


        private void DeleteAgentBtn(object sender, EventArgs e)
        {
            if (DataGridView.SelectedRows.Count == 1)
            {
                int agentID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["AgentID"].Value);

                if (MessageBox.Show("Bu agenti silmək istədiyinizə əminsiniz?", "Təsdiq", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        connection.Open();

                        SqlCommand command = new SqlCommand("UPDATE Agent SET DeletedDate = @DeletedDate WHERE AgentId = @AgentId", connection);
                        command.Parameters.AddWithValue("@DeletedDate", DateTime.UtcNow.AddHours(+4)); 
                        command.Parameters.AddWithValue("@AgentId", agentID);

                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Agent uğurla silindi.", "Uğurlu Əməliyyat", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            BindData();
                        }
                        else
                        {
                            MessageBox.Show("Agent silinmədi.", "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Xəta baş verdi: " + ex.Message, "Xəta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Silinmə əməliyyatı ləğv edildi.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            ClearInputFields();
        }

        private void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow selectedRow = DataGridView.Rows[e.RowIndex];
                textBox2.Text = selectedRow.Cells["Name"].Value.ToString();
            }
        }

        private void ExportToExcelBtn(object sender, EventArgs e)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Agent Data");

                for (int i = 0; i < DataGridView.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = DataGridView.Columns[i].HeaderText;
                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                }

                for (int row = 0; row < DataGridView.Rows.Count; row++)
                {
                    for (int col = 0; col < DataGridView.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = DataGridView.Rows[row].Cells[col].Value.ToString();
                    }
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.FileName = "AgentData.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        excelPackage.SaveAs(excelFile);
                        MessageBox.Show("Excell faylı kimi yükləndi", "Export Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (connection.State != ConnectionState.Open)
            {
                connection.Open();
            }

            string searchText = textBox1.Text.Trim();

            string query;

            if (string.IsNullOrWhiteSpace(searchText))
            {
                query = "SELECT * FROM Agent";
            }
            else
            {
                query = "SELECT * FROM Agent WHERE Name LIKE @SearchText";
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


        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                BindData();
            }

            if (e.KeyChar == (char)13)
                button5.PerformClick();
        }
    }
}
