using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Maliyye
{

    public partial class Agent : Form
    {
        public Agent()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Agent_Load(object sender, EventArgs e)
        {
            BindData();
        }

        SqlConnection connection = new SqlConnection("Data Source=WIN-AR9UPHEFUDQ\\SQLEXPRESS;Initial Catalog=Agent;Integrated Security=True");

        void BindData()
        {
            SqlCommand command = new SqlCommand("select * from Agent", connection);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            connection.Open();
            SqlCommand command = new SqlCommand("INSERT INTO Agent VALUES (@Value)", connection);
            command.Parameters.AddWithValue("@Value", textBox2.Text);
            command.ExecuteNonQuery();
            MessageBox.Show("Inserted successfully");
            connection.Close();
            BindData();


        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                SqlCommand command = new SqlCommand("UPDATE Agent SET Name = @Value WHERE AgentId = @ID", connection);
                command.Parameters.AddWithValue("@Value", textBox2.Text);
                command.Parameters.AddWithValue("@ID", int.Parse(textBox1.Text)); // Replace this with the appropriate record ID

                int rowsAffected = command.ExecuteNonQuery();

                if (rowsAffected > 0)
                {
                    MessageBox.Show("Updated successfully");
                    BindData();
                }
                else
                {
                    MessageBox.Show("Update operation failed");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            connection.Close();
            BindData();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                SqlCommand command = new SqlCommand("DELETE FROM Agent WHERE AgentId = @ID", connection);
                command.Parameters.AddWithValue("@ID", int.Parse(textBox1.Text)); // Replace this with the appropriate record ID

                int rowsAffected = command.ExecuteNonQuery();

                if (rowsAffected > 0)
                {
                    MessageBox.Show("Deleted successfully");
                    BindData();
                }
                else
                {
                    MessageBox.Show("Delete operation failed");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }

        }

 

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }

}
