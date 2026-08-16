using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Security.Cryptography;

namespace ImportData.Core
{
    /// <summary>
    /// Lớp AppConfig: Quản lý cấu hình kết nối SQL và thư mục giám sát.
    /// Thiết kế chuẩn Portable: Đảm bảo chạy mượt mà trên MỌI MÁY TÍNH khác nhau.
    /// </summary>
    public class AppConfig
    {
        private const string DefaultFolderName = "task";
        public const string DefaultConnectionString = "Server=dbserver.hycap.co.kr,5398;Database=SmartFactoryV2;User ID=vinaadmin;Password=vina1234%6&8;Encrypt=False;TrustServerCertificate=True;";

        public string ConnectionString { get; set; } = DefaultConnectionString;
        public string BaseFolder { get; set; }
        public int ScanIntervalSeconds { get; set; } = 600;
        public int HealthCheckTimeoutSeconds { get; set; } = 5;

        private string _configFilePath = "";

        public AppConfig() 
        {
            ConnectionString = DefaultConnectionString;
            BaseFolder = @"C:\task";
        }

        public void Load(Action<string>? logger = null)
        {
            try
            {
                string runDir = AppDomain.CurrentDomain.BaseDirectory;
                string settingsPath = Path.Combine(runDir, "appsettings.json");
                
                // Ưu tiên đọc file appsettings.json ngay tại thư mục chứa file .exe
                if (!File.Exists(settingsPath))
                {
                    string sourceDir = Path.GetFullPath(Path.Combine(runDir, @"..\..\..\"));
                    string sourceSettingsPath = Path.Combine(sourceDir, "appsettings.json");
                    if (File.Exists(sourceSettingsPath)) settingsPath = sourceSettingsPath;
                }

                _configFilePath = settingsPath;

                if (File.Exists(_configFilePath))
                {
                    string json;
                    using (var fs = new FileStream(_configFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var sr = new StreamReader(fs))
                    {
                        json = sr.ReadToEnd();
                    }

                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        using (var doc = JsonDocument.Parse(json))
                        {
                            var root = doc.RootElement;
                            
                            if (root.TryGetProperty("ConnectionStrings", out var connSection))
                            {
                                string? rawConn = connSection.GetProperty("DefaultConnection").GetString();
                                if (!string.IsNullOrWhiteSpace(rawConn))
                                {
                                    if (rawConn.StartsWith("ENC:"))
                                    {
                                        string? decrypted = Decrypt(rawConn);
                                        // Nếu giải mã thất bại (do copy sang máy khác), tự động Fallback về chuỗi mặc định
                                        ConnectionString = !string.IsNullOrEmpty(decrypted) ? decrypted : DefaultConnectionString;
                                    }
                                    else
                                    {
                                        ConnectionString = rawConn;
                                    }
                                }
                            }

                            if (root.TryGetProperty("FolderSettings", out var folderSection))
                            {
                                BaseFolder = folderSection.GetProperty("BaseFolder").GetString() ?? BaseFolder;
                            }

                            if (root.TryGetProperty("SyncSettings", out var syncSection))
                            {
                                ScanIntervalSeconds = syncSection.GetProperty("ScanIntervalSeconds").GetInt32();
                            }

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
                ConnectionString = DefaultConnectionString;
                logger?.Invoke($"[CẢNH BÁO] Dùng cấu hình mặc định: {ex.Message}");
            }
        }

        public void Save()
        {
            try
            {
                if (string.IsNullOrEmpty(_configFilePath))
                {
                    _configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
                }

                var configData = new
                {
                    ConnectionStrings = new { DefaultConnection = ConnectionString }, // Lưu dạng rõ ràng để máy khác copy sang dùng được ngay
                    FolderSettings = new { BaseFolder = BaseFolder },
                    SyncSettings = new { ScanIntervalSeconds = ScanIntervalSeconds },
                    HealthCheckSettings = new { ConnectionTimeoutSeconds = HealthCheckTimeoutSeconds }
                };

                string json = JsonSerializer.Serialize(configData, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_configFilePath, json);
            }
            catch { }
        }

        private string? Encrypt(string plainText)
        {
            if (string.IsNullOrEmpty(plainText) || plainText.StartsWith("ENC:")) return plainText;
            try {
                byte[] data = Encoding.UTF8.GetBytes(plainText);
                byte[] encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
                return "ENC:" + Convert.ToBase64String(encrypted);
            } catch { return plainText; }
        }

        private string? Decrypt(string cipherText)
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
