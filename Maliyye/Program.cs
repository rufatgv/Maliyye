using System;
using Maliyye.Forms;
using System.Windows.Forms;
using System.Configuration;

namespace Maliyye
{
    static class Program
    {
        public static string dataSource = ConfigurationManager.ConnectionStrings["Maliyye.Properties.Settings.AgentConnectionString"].ConnectionString;
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());

            
        }
    }
}
