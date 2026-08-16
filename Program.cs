using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ImportData 
{
    internal static class Program
    {
        private static Mutex? _mutex;

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool AllocConsole();

        [STAThread] 
        static void Main(string[] args) 
        {
            bool isDebugMode = args.Length > 0 && (args[0].Equals("--console", StringComparison.OrdinalIgnoreCase) || 
                                                   args[0].Equals("-debug", StringComparison.OrdinalIgnoreCase) || 
                                                   args[0].Equals("/debug", StringComparison.OrdinalIgnoreCase));

            if (isDebugMode)
            {
                AllocConsole();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("=================================================");
                Console.WriteLine("    AUTO IMPORT DATA - CHE DO CHANDOAN CHUYEN SAU");
                Console.WriteLine("=================================================");
                Console.ResetColor();
                Console.WriteLine($"[1/5] Thoi gian: {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
                Console.WriteLine($"[1/5] Thu muc thuc thi: {AppDomain.CurrentDomain.BaseDirectory}");
            }

            // 1. Chuyển Microsoft.Data.SqlClient sang pure C# Managed Networking
            try
            {
                if (isDebugMode) Console.WriteLine("[2/5] Dang thiet lap Pure C# Managed Networking cho SQL...");
                AppContext.SetSwitch("Switch.Microsoft.Data.SqlClient.UseManagedNetworkingOnWindows", true);
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (isDebugMode) Console.WriteLine("[2/5] Thiet lap Managed Networking: THANH CONG.");
            }
            catch (Exception ex)
            {
                if (isDebugMode)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"[2/5] Lỗi Managed Networking: {ex.Message}");
                    Console.ResetColor();
                }
            }

            // 2. Bắt toàn bộ lỗi khởi động chưa được xử lý
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (sender, e) => HandleException(e.Exception, isDebugMode);
            AppDomain.CurrentDomain.UnhandledException += (sender, e) => HandleException(e.ExceptionObject as Exception, isDebugMode);

            // 3. Khởi tạo DPI tương thích với MỌI phiên bản Windows
            try
            {
                if (isDebugMode) Console.WriteLine("[3/5] Dang khoi tao giao dien DPI Windows Forms...");
                ApplicationConfiguration.Initialize();
                if (isDebugMode) Console.WriteLine("[3/5] Khoi tao DPI: THANH CONG.");
            }
            catch (Exception ex)
            {
                if (isDebugMode) Console.WriteLine($"[3/5] Canh bao DPI fallback: {ex.Message}");
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
            }

            // 4. Kiểm tra ứng dụng duy nhất (Single Instance Mutex an toàn)
            bool createdNew = false;
            try
            {
                const string mutexName = "Global\\ImportData_Vinatech_AutoImport_SingleInstance_Mutex";
                _mutex = new Mutex(true, mutexName, out createdNew);
            }
            catch (Exception ex)
            {
                if (isDebugMode) Console.WriteLine($"[4/5] Canh bao Mutex: {ex.Message}");
                createdNew = true;
            }

            if (!createdNew)
            {
                string runningMsg = "Ứng dụng Auto Import Data ĐANG CHẠY NGẦM TRÊN MÁY TÍNH!\n\nVui lòng kiểm tra khay đồng hồ (góc dưới bên phải màn hình) và nhấp đúp vào biểu tượng ứng dụng để hiển thị lại cửa sổ.";
                if (isDebugMode)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"[4/5] THONG BAO: App dang chay ngam!");
                    Console.ResetColor();
                    Console.WriteLine("Nhan ENTER de thoat...");
                    Console.ReadLine();
                }
                MessageBox.Show(runningMsg, "Thông Báo Ứng Dụng Đã Mở", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                if (isDebugMode) Console.WriteLine("[5/5] Dang khoi tao Form1 va hien thi giao dien...");
                Form1 mainForm = new Form1();
                if (isDebugMode) Console.WriteLine("[5/5] Form1 da khoi tao thanh cong. Dang chay Application.Run()...");
                Application.Run(mainForm);
            }
            catch (Exception ex)
            {
                HandleException(ex, isDebugMode);
            }
            finally
            {
                try { _mutex?.ReleaseMutex(); } catch { }
            }
        }

        private static void HandleException(Exception? ex, bool isDebugMode = false)
        {
            if (ex == null) return;
            string msg = $"[ĐÃ XẢY RA LỖI KHỞI ĐỘNG HỆ THỐNG]\n\nChi tiết lỗi: {ex.Message}\n\nVị trí phát sinh:\n{ex.StackTrace}";
            
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "startup_error.txt");
                File.WriteAllText(logPath, msg);
            }
            catch { }

            if (isDebugMode)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\n=================================================");
                Console.WriteLine("             PHAT HIEN LOI nghiem trong");
                Console.WriteLine("=================================================");
                Console.WriteLine(msg);
                Console.ResetColor();
                Console.WriteLine("\nNhan phim ENTER de thoat...");
                Console.ReadLine();
            }

            MessageBox.Show(msg, "Lỗi Khởi Động ImportData", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}