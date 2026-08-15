using System;
using System.Windows.Forms;

namespace ImportData 
{
    /// <summary>
    /// Lớp Program: Gốc khởi đầu cho một chương trình chuẩn trên C# Windows.
    /// Quản lý việc cấu hình khung và gọi tới Form giao diện ban đầu phần mềm.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Thuộc tính [STAThread] (Căn hộ mô hình Luồng Phân Đơn Single-Threaded Apartment): 
        /// Đây là điểm bắt buộc định dạng dành cho các phần mềm liên đới với nền tảng đồ họa Windows.
        /// Bảo đảm mọi lệnh kích phím UI đều rơi được về một hướng mạch ổn định an toàn từ O.S, không rơi vào lỗi xung đột Thread.
        /// </summary>
        [STAThread] 
        static void Main() 
        {
            // ApplicationConfiguration.Initialize() để tự động xác nhận thông số kỹ thuật màn hình máy (System DPI).
            // Cho phần mềm khi trải qua màn hình LED, hay máy 4K thì phông nền màn lưới C# Form vẫn siêu sắc nét không hư nhòe hạt (DPI High Config).
            ApplicationConfiguration.Initialize();

            // Hàm Run(): Yêu cầu Hệ điều hành vẽ hộp khối Cửa Sổ Máy tên Form1() (Giao kiện chính của bạn) lên Màn Desktop. 
            // Khi chạy qua hàm này ứng dụng lập tức chìm vào Chu Trình 'Message Loop' Của Máy nhằm Lắng Nghe Khách tương tác click nút thao tác cho tới lúc Khách chủ động ngắt phần mềm!.
            Application.Run(new Form1());
        }
    }
}