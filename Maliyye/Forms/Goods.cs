using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Windows.Forms;
using Maliyye.AppCode.Extensions;
using System.Configuration;
using OfficeOpenXml;
using System.IO;

namespace Maliyye.Forms
{
    public partial class Goods : Form
    {
        SqlConnection connection = new SqlConnection(Program.dataSource);
        public Goods()
        {
            InitializeComponent();
            DataGridView.InitDefault();
            DataGridView.CellClick += DataGridView_CellClick;
            DataGridView.ReadOnly = true;
            this.KeyPreview = true;
            this.KeyDown += Agent_KeyDown;
        }

        private void Goods_Load(object sender, EventArgs e)
        {
            BindData();
        }

        private void Agent_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        void BindData()
        {
            SqlCommand command = new SqlCommand("select * from Goods", connection);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            DataGridView.DataSource = dt;
            DataGridView.Columns["GoodsID"].Visible = false;
            DataGridView.Columns[0].HeaderText = "Malın adı";
            DataGridView.Columns[1].HeaderText = "Qiyməti";
            DataGridView.Columns[3].HeaderText = "Silinmə tarixi";
        }

        private void ClearInputFields()
        {
            textBox3.Clear();
            textBox2.Clear();
        }

        private void CreateGoodsBtn(object sender, EventArgs e)
        {
            string name = textBox2.Text;
            string priceText = textBox3.Text;

            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(priceText))
            {
                MessageBox.Show("Malın adını və qiymətini qeyd edin");
                return;
            }

            if (!decimal.TryParse(priceText, out decimal price))
            {
                MessageBox.Show("Qiymət dəyəri düzgün bir nömrə deyil. Zəhmət olmasa bir rəqəm daxil edin.");
                return;
            }

            try
            {
                connection.Open();
                SqlCommand command = new SqlCommand("INSERT INTO Goods (Name, Price, DeletedDate) VALUES (@Name, @Price, @DeletedDate)", connection);
                command.Parameters.AddWithValue("@Name", name);
                command.Parameters.AddWithValue("@Price", price);
                command.Parameters.AddWithValue("@DeletedDate", DBNull.Value);
                command.ExecuteNonQuery();
                MessageBox.Show("Əlavə edildi");
                BindData();


            }
            catch (Exception ex)
            {
                MessageBox.Show("Bir səhv baş verdi: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
            ClearInputFields();
        }

        private void EditGoodsBtn(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text) || string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("Malın adını və ya qiymətini boş qoyma", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            int goodsID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["GoodsID"].Value);
            DateTime? deletedDate = GetGoodsDeletedDate(goodsID);


            if (deletedDate == null)
            if (deletedDate == null)
            {
                if (MessageBox.Show("Redaktə etmək istədiyinizdən əminsiniz?", "Təsdiq", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        connection.Open();

                        SqlCommand command = new SqlCommand("UPDATE Goods SET Name = @Name, Price = @Price, DeletedDate = @DeletedDate WHERE GoodsID = @ID", connection);
                        command.Parameters.AddWithValue("@Name", textBox2.Text);
                        command.Parameters.AddWithValue("@Price", decimal.Parse(textBox3.Text));
                        command.Parameters.AddWithValue("@ID", goodsID);
                        command.Parameters.AddWithValue("@DeletedDate", DBNull.Value);

                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Mal yeniləndi", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            BindData();
                        }
                        else
                        {
                            MessageBox.Show("Mal yenilənmədi", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Səhv baş verdi: " + ex.Message, "Səhv", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Redaktə əməliyyatı ləğv edildi.", "Ləğv", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Silinən agenti redaktə edə bilməzsiniz.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            ClearInputFields();
        }

        private void DeleteGoodsBtn(object sender, EventArgs e)
        {
            if (DataGridView.SelectedRows.Count == 1)
            {
                int goodsID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["GoodsID"].Value);

                if (MessageBox.Show("Bu malı silmək istədiyinizdən əminsiniz?", "Təsdiq", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        connection.Open();

                        SqlCommand command = new SqlCommand("Update Goods SET DeletedDate = @DeletedDate WHERE GoodsID = @ID", connection);
                        command.Parameters.AddWithValue("@ID", goodsID);
                        command.Parameters.AddWithValue("@DeletedDate", DateTime.UtcNow.AddHours(+4));

                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Mal uğurla silindi.", "Uğurlu", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            BindData();
                        }
                        else
                        {
                            MessageBox.Show("Mal silinmədi.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Səhv baş verdi: " + ex.Message, "Səhv", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Silinmə əməliyyatı ləğv edildi.", "Ləğv", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("Xahiş edirik, silmək üçün DataGridView-dan bir sətr seçin.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            ClearInputFields();
        }

        private DateTime? GetGoodsDeletedDate(int goodsID)
        {
            connection.Open();
            SqlCommand command = new SqlCommand("SELECT DeletedDate FROM Goods WHERE GoodsID = @GoodsID", connection);
            command.Parameters.AddWithValue("@GoodsID", goodsID);
            object result = command.ExecuteScalar();
            connection.Close();
            if (result != null && result != DBNull.Value)
            {
                return Convert.ToDateTime (result);
            }
            return null;
        }

        private void SearchGoods(string searchValue)
        {
            foreach (DataGridViewRow row in DataGridView.Rows)
            {
                DataGridViewCell nameCell = row.Cells["Name"];

                if (nameCell.Value != null && nameCell.Value.ToString().Contains(searchValue))
                {
                    DataGridView.ClearSelection();
                    row.Selected = true;

                    if (row.Index >= 0 && row.Index < DataGridView.Rows.Count)
                    {
                        DataGridView.FirstDisplayedScrollingRowIndex = row.Index;
                    }

                    textBox2.Text = row.Cells["Name"].Value.ToString();
                    textBox3.Text = row.Cells["Price"].Value.ToString();

                    return;
                }
            }

            MessageBox.Show("No matching records found.");
        }

        private void SearchBtn_Click(object sender, EventArgs e)
        {
            string searchValue = textBox4.Text.Trim();
            if (!string.IsNullOrEmpty(searchValue))
            {
                SearchGoods(searchValue);
            }
            else
            {
                MessageBox.Show("Please enter a search value.");
            }
        }

        private void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow selectedRow = DataGridView.Rows[e.RowIndex];
                string name = selectedRow.Cells["Name"].Value.ToString();
                decimal price = Convert.ToDecimal(selectedRow.Cells["Price"].Value);
                textBox2.Text = name;
                textBox3.Text = price.ToString();
            }
        }

        private void Search_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Mallar");

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
                    saveFileDialog.FileName = "Mallar.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        excelPackage.SaveAs(excelFile);
                        MessageBox.Show("Excell faylı kimi yükləndi", "Export Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                BindData();
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

            string searchText = textBox1.Text.Trim();

            string query;

            if (string.IsNullOrWhiteSpace(searchText))
            {
                query = "SELECT * FROM Goods";
            }
            else
            {
                query = "SELECT * FROM Goods WHERE Name LIKE @SearchText";
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
