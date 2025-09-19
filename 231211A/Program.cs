using System;
using System.Windows.Forms;

namespace _231211A
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // 高 DPI (每螢幕) 支援，避免跨螢幕縮放模糊或跳動
            try { Application.SetHighDpiMode(HighDpiMode.PerMonitorV2); } catch { }

            // 預設 WinForms 初始化 (啟用視覺樣式 / 預設字型等)
            ApplicationConfiguration.Initialize();

            Application.Run(new Form1());
        }
    }
}