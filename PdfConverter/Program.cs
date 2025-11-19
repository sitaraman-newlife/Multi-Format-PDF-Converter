using System;
using System.Windows.Forms;

namespace PdfConverter
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Enable visual styles for modern Windows Forms appearance
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // Initialize and run the main application form
            Application.Run(new Forms.MainForm());
        }
    }
}
