using System;
using System.Windows.Forms;

namespace Maliyye.Forms
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void AgentBtn_Click(object sender, EventArgs e)
        {
            Agent webForm = new Agent();
            webForm.ShowDialog();
        }

        private void GoodsBtn_Click(object sender, EventArgs e)
        {
            Goods webForm = new Goods();
            webForm.ShowDialog();
        }

        private void SalesBtn_Click(object sender, EventArgs e)
        {
            Sales webForm = new Sales();
            webForm.ShowDialog();
        }

        private void PaymentsBtn_Click(object sender, EventArgs e)
        {
            using (var frm = new Payments())
            {
                frm.ShowDialog();
            }
        }

        private void PaymentsModule_Click(object sender, EventArgs e)
        {
            PaymentsModule webForm = new PaymentsModule();
            webForm.ShowDialog();
        }

        private void Invoice_Click(object sender, EventArgs e)
        {
            Invoice webForm = new Invoice();
            webForm.ShowDialog();
        }

        private void PaymentsToAgent_Click(object sender, EventArgs e)
        {
            PaymentsToAgent webForm = new PaymentsToAgent();
            webForm.ShowDialog();
        }

        private void Report_Click(object sender, EventArgs e)
        {
            Hesabat webForm = new Hesabat();
            webForm.ShowDialog();
        }
    }
}
