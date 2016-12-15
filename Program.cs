using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;
using MetroFramework;

namespace WindowsFormsApplication1
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>

        public static Process PriorProcess()
        {
            Process curr = Process.GetCurrentProcess();
            Process[] procs = Process.GetProcessesByName(curr.ProcessName);
            foreach (Process p in procs)
            {
                if ((p.Id != curr.Id) &&
                (p.MainModule.FileName == curr.MainModule.FileName))
                    return p;
            }
            return null;
        }
        
        [STAThread]
        static void Main()
        {
            if (PriorProcess() != null)
            {
                MessageBox.Show("Another instance of Mi-LAW is already running");
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Login Login = new Login();
            if (Login.ShowDialog() == DialogResult.OK)
            {
                Application.Run(new MainForm());
            }
            else
            {
                Application.Exit();
            }
        }
    }
}
