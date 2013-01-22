using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Pace
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // only allow one instance of Pace to be running
            System.Diagnostics.Process[] p1 = System.Diagnostics.Process.GetProcessesByName("pace");

            if (p1.Length > 1)
            {
                Application.Exit();
            }
            else
            {
                Application.Run(new Form1());
            }
        }
    }
}