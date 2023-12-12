using Maliyye.AppCode.Extensions;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace Maliyye.Forms
{
    public partial class Hesabat : Form
    {

        SqlConnection connection = new SqlConnection(Program.dataSource);
        public Hesabat()
        {
            InitializeComponent();
            dataGridView1.InitDefault();
            BindData();
            dataGridView1.Columns[0].HeaderText = "Agent adı";
            dataGridView1.Columns[1].HeaderText = "Mal adı";
            dataGridView1.Columns[2].HeaderText = "Say";
            dataGridView1.Columns[3].HeaderText = "Bonus";
            dataGridView1.Columns[4].HeaderText = "Borc";
            dataGridView1.Columns[5].HeaderText = "Son borc";
            dataGridView1.Columns[6].HeaderText = "Göndərilən qaimə";
            dataGridView1.Columns[7].HeaderText = "Göndərilməyən qaimə";
            dataGridView1.Columns[8].HeaderText = "Cəm";
            dataGridView1.Columns[9].HeaderText = "Ödənilən məbləğ";
            dataGridView1.ReadOnly = true;
            this.KeyPreview = true;
            this.KeyDown += Agent_KeyDown;
        }

        void BindData()
        {
            SqlCommand command = new SqlCommand("SELECT * FROM vWReport", connection);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void Agent_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (connection.State != ConnectionState.Open)
            {
                connection.Open();
            }

            string searchText = textBox1.Text.Trim();

            string query;

            if (string.IsNullOrWhiteSpace(searchText))
            {
                query = "SELECT * FROM vWReport";
            }
            else
            {
                query = "SELECT * FROM vWReport WHERE AgentName LIKE @SearchText OR GoodsName LIKE @SearchText";
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

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                BindData();
            }

            if (e.KeyChar == (char)13)
                button1.PerformClick();

        }

        private void ExportToExcelBtn_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Hesabat");

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;

                    // Set the header cell color (change Color.Red to your desired color)
                    worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);

                    worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                }

                for (int row = 0; row < dataGridView1.Rows.Count; row++)
                {
                    for (int col = 0; col < dataGridView1.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dataGridView1.Rows[row].Cells[col].Value.ToString();
                    }
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.FileName = "Hesabat.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                        excelPackage.SaveAs(excelFile);
                        MessageBox.Show("Excel faylı kimi yükləndi", "Export Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }


    }
}

