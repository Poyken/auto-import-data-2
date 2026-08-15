using System.Text.Json;
using System.Security.Cryptography;
using System.Text;

namespace ImportData.Core
{
    /// <summary>
    /// Lớp AppConfig: Quản lý các cấu hình và lưu trữ thông tin hệ thống.
    /// Giúp ứng dụng biết "Máy chủ SQL nào?" và "Thư mục nào cần quét?".
    /// </summary>
    public class AppConfig
    {
        // Tên thư mục mặc định trên Desktop để canh chừng file mới.
        private const string DefaultFolderName = "task";

        // CÁC THUỘC TÍNH CẤU HÌNH:
        public string ConnectionString { get; set; }     // Chuỗi kết nối SQL Server.
        public string BaseFolder { get; set; }           // Đường dẫn thư mục máy đo.
        public int ScanIntervalSeconds { get; set; } = 600; // Tần suất quét định kỳ (giây).
        public int HealthCheckTimeoutSeconds { get; set; } = 5; // Thời gian chờ SQL tối đa.

        private string _configFilePath; // Lưu đường dẫn file appsettings.json đã nạp.

        // HÀM KHỞI TẠO (Constructor): Chạy đầu tiên khi tạo mới đối tượng AppConfig.
        public AppConfig() 
        {
            // Thiết lập giá trị mặc định ban đầu.
            ConnectionString = "";
            BaseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), DefaultFolderName); 
        }

        /// <summary>
        /// Nạp cấu hình từ file appsettings.json lên bộ nhớ RAM.
        /// </summary>
        public void Load(Action<string> logger = null)
        {
            try
            {
                // Tìm kiếm file appsettings.json ở thư mục chạy hoặc thư mục source code.
                string runDir = AppDomain.CurrentDomain.BaseDirectory; 
                string sourceDir = Path.GetFullPath(Path.Combine(runDir, @"..\..\..\")); 
                
                string settingsPath = Path.Combine(runDir, "appsettings.json"); 
                string sourceSettingsPath = Path.Combine(sourceDir, "appsettings.json"); 

                // Chọn file nào tồn tại (Ưu tiên thư mục Source khi đang Code).
                _configFilePath = File.Exists(sourceSettingsPath) ? sourceSettingsPath : settingsPath; 

                if (File.Exists(_configFilePath))
                {
                    string json;
                    // Mở file và đọc toàn bộ nội dung (hỗ trợ đọc khi file đang mở).
                    using (var fs = new FileStream(_configFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) 
                    using (var sr = new StreamReader(fs)) 
                    {
                        json = sr.ReadToEnd();
                    }

                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        // Phân tích chuỗi JSON thành các thuộc tính C#.
                        using (var doc = JsonDocument.Parse(json)) 
                        {
                            var root = doc.RootElement; 
                            
                            // 1. Nạp Chuỗi Kết Nối (Nếu không ép dùng chuỗi fix cứng).
                            if (root.TryGetProperty("ConnectionStrings", out var connSection))
                            {
                                string rawConn = connSection.GetProperty("DefaultConnection").GetString();
                                ConnectionString = Decrypt(rawConn) ?? rawConn;
                            }

                            // 2. Nạp Đường Dẫn Thư Mục.
                            if (root.TryGetProperty("FolderSettings", out var folderSection))
                            {
                                BaseFolder = folderSection.GetProperty("BaseFolder").GetString() ?? BaseFolder;
                            }

                            // 3. Nạp Tần Suất Quét.
                            if (root.TryGetProperty("SyncSettings", out var syncSection))
                            {
                                ScanIntervalSeconds = syncSection.GetProperty("ScanIntervalSeconds").GetInt32();
                            }

                            // 4. Nạp Thời Gian Chờ Kết Nối.
                            if (root.TryGetProperty("HealthCheckSettings", out var healthSection))
                            {
                                HealthCheckTimeoutSeconds = healthSection.GetProperty("ConnectionTimeoutSeconds").GetInt32();
                            }
                        }
                    }
                    logger?.Invoke($"[CẤU HÌNH] Đã nạp thành công từ: {Path.GetFileName(_configFilePath)}");
                }
            }
            catch (Exception ex)
            {
                // Nếu lỗi (file hỏng, sai định dạng), in cảnh báo và dùng mặc định.
                logger?.Invoke($"[CẢNH BÁO] Không nạp được cấu hình, dùng mặc định: {ex.Message}");
            }
        }

        /// <summary>
        /// Lưu các thông số hiện tại xuống file appsettings.json.
        /// </summary>
        public void Save()
        {
            try
            {
                if (string.IsNullOrEmpty(_configFilePath)) return;

                // Tạo đối tượng JSON để lưu trữ.
                var configData = new
                {
                    ConnectionStrings = new { DefaultConnection = Encrypt(ConnectionString) },
                    FolderSettings = new { BaseFolder = BaseFolder },
                    SyncSettings = new { ScanIntervalSeconds = ScanIntervalSeconds },
                    HealthCheckSettings = new { ConnectionTimeoutSeconds = HealthCheckTimeoutSeconds }
                };

                // Chuyển đối tượng C# thành định dạng văn bản JSON đẹp mắt.
                string json = JsonSerializer.Serialize(configData, new JsonSerializerOptions { WriteIndented = true });
                // Ghi đè vào file trên ổ cứng.
                File.WriteAllText(_configFilePath, json);
            }
            catch (Exception)
            {
                // Bỏ qua lỗi nếu không thể lưu (ví dụ file đang bị khóa bởi trình khác).
            }
        }

        // --- CÁC HÀM HỖ TRỢ MÃ HÓA (Dùng DPAPI của Windows) ---
        private string Encrypt(string plainText)
        {
            if (string.IsNullOrEmpty(plainText) || plainText.StartsWith("ENC:")) return plainText;
            try {
                byte[] data = Encoding.UTF8.GetBytes(plainText);
                byte[] encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
                return "ENC:" + Convert.ToBase64String(encrypted);
            } catch { return plainText; }
        }

        private string Decrypt(string cipherText)
        {
            if (string.IsNullOrEmpty(cipherText) || !cipherText.StartsWith("ENC:")) return null;
            try {
                byte[] data = Convert.FromBase64String(cipherText.Substring(4));
                byte[] decrypted = ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(decrypted);
            } catch { return null; }
        }
    }
}
