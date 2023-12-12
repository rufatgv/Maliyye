using Maliyye.AppCode.Extensions;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace Maliyye.Forms
{
    public partial class Payments : Form
    {
        SqlConnection connection = new SqlConnection(Program.dataSource);
        public Payments()
        {
            InitializeComponent();
            DataGridView.InitDefault();
            DataGridView.CellClick += DataGridView_CellClick;
            DataGridView.ReadOnly = true;
            this.KeyPreview = true;
            this.KeyDown += Agent_KeyDown;
        }

        private void Payments_Load(object sender, EventArgs e)
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

        private void ClearInputFields()
        {
            textBox2.Clear();
        }
        void BindData()
        {
            SqlCommand command = new SqlCommand("select * from Payments", connection);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            DataGridView.DataSource = dt;
            DataGridView.Columns["PaymentID"].Visible = false;
            DataGridView.Columns[1].HeaderText = "Ödənişin növü";
            DataGridView.Columns[2].HeaderText = "Silinmə tarixi";
        }

        private void CreatePaymentBtn(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                MessageBox.Show("Ödəniş növü daxil edin", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            connection.Open();
            SqlCommand command = new SqlCommand("INSERT INTO Payments (DeletedDate,PaymentType) VALUES (@DeletedDate,@PaymentType)", connection);
            command.Parameters.AddWithValue("@PaymentType", textBox2.Text);
            command.Parameters.AddWithValue("@DeletedDate", DBNull.Value);

            command.ExecuteNonQuery();
            MessageBox.Show("Əlavə edildi");
            connection.Close();
            BindData();
            ClearInputFields();
        }

        private void EditPaymentBtn(object sender, EventArgs e)
        {
            if (DataGridView.SelectedRows.Count == 1)
            {
                int paymentID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["PaymentID"].Value);
                DateTime? deletedDate = GetPaymentsDeletedDate(paymentID);

                if (deletedDate == null)
                {
                    if (MessageBox.Show("Siz bu ödənişi yeniləmək istədiyinizdən əminsiniz?", "Təsdiq", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            connection.Open();

                            string sql = "UPDATE Payments SET PaymentType = @PaymentType ,DeletedDate = @DeletedDate WHERE PaymentID = @PaymentID";

                            using (SqlCommand cmd = new SqlCommand(sql, connection))
                            {
                                if (string.IsNullOrWhiteSpace(textBox2.Text))
                                {
                                    MessageBox.Show("Ödəniş növü daxil edin", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    return;
                                }

                                cmd.Parameters.AddWithValue("@PaymentType", textBox2.Text);
                                cmd.Parameters.AddWithValue("@PaymentID", paymentID);
                                cmd.Parameters.AddWithValue("@DeletedDate", DBNull.Value);

                                int rowsAffected = cmd.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    MessageBox.Show("Ödəniş uğurla yeniləndi.");
                                    BindData();
                                }
                                else
                                {
                                    MessageBox.Show("Ödəniş yenilənmədi.");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Səhv baş verdi: " + ex.Message);
                        }
                        finally
                        {
                            connection.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Yeniləmə əməliyyatı ləğv edildi.");
                    }
                }
                else
                {
                    MessageBox.Show("Silinən ödəniş növünü redaktə edə bilməzsiniz.", "Xəbərdarlıq", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                ClearInputFields();
            }
        }

        private DateTime? GetPaymentsDeletedDate(int paymentID)
        {
            connection.Open();
            SqlCommand command = new SqlCommand("SELECT DeletedDate FROM Payments WHERE PaymentID = @PaymentID", connection);
            command.Parameters.AddWithValue("@PaymentID", paymentID);
            object result = command.ExecuteScalar();
            connection.Close();
            if (result != null && result != DBNull.Value)
            {
                return Convert.ToDateTime(result);
            }
            return null;
        }

        private void DeletePaymentBtn(object sender, EventArgs e)
        {
            if (DataGridView.SelectedRows.Count == 1)
            {
                int paymentID = Convert.ToInt32(DataGridView.SelectedRows[0].Cells["PaymentID"].Value);

                if (MessageBox.Show("Siz bu ödənişi silmək istədiyinizdən əminsiniz?", "Təsdiq", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    try
                    {
                        connection.Open();
                        SqlCommand command = new SqlCommand("Update Payments SET DeletedDate = @DeletedDate WHERE PaymentID = @ID", connection);
                        command.Parameters.AddWithValue("@ID", paymentID);
                        command.Parameters.AddWithValue("@DeletedDate", DateTime.UtcNow.AddHours(+4));

                        int rowsAffected = command.ExecuteNonQuery();

                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Silindi");
                            BindData();
                        }
                        else
                        {
                            MessageBox.Show("Silinmə əməliyyatı uğurlu olmadı");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Səhv baş verdi: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
                else
                {
                    MessageBox.Show("Silinmə əməliyyatı ləğv edildi.");
                }
            }
            ClearInputFields();
        }
        private void SearchPayment(string searchValue)
        {
            foreach (DataGridViewRow row in DataGridView.Rows)
            {
                DataGridViewCell nameCell = row.Cells["PaymentType"];

                if (nameCell.Value != null && nameCell.Value.ToString().Contains(searchValue))
                {
                    DataGridView.ClearSelection();
                    row.Selected = true;

                    if (row.Index >= 0 && row.Index < DataGridView.Rows.Count)
                    {
                        DataGridView.FirstDisplayedScrollingRowIndex = row.Index;
                    }

                    textBox3.Text = row.Cells["PaymentType"].Value.ToString();

                    return;
                }
            }
            MessageBox.Show("Tapilmadi");
        }

        private void SearchBtn_Click(object sender, EventArgs e)
        {
            string searchValue = textBox3.Text.Trim();
            if (!string.IsNullOrEmpty(searchValue))
            {
                SearchPayment(searchValue);
            }
            else
            {
                MessageBox.Show("Xahiş edirəm axtarış dəyəri daxil edin.");
            }
        }

        private void DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow selectedRow = DataGridView.Rows[e.RowIndex];

                string paymentType = selectedRow.Cells["PaymentType"].Value.ToString();
                textBox2.Text = paymentType;
            }
        }

        private void Search_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Ödəniş növü");

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
                    saveFileDialog.FileName = "Ödəniş növü.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        excelPackage.SaveAs(excelFile);
                        MessageBox.Show("Excell faylı kimi yükləndi", "Export Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }
    }
}
