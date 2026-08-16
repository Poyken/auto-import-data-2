using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Concurrent;
using ImportData.Core;
using ImportData.Services;

namespace ImportData
{
    /// <summary>
    /// Giao diện chính Form1: Trung tâm điều phối của ứng dụng Auto Import.
    /// Quản lý việc theo dõi thư mục thời gian thực, đồng bộ dữ liệu và hiển thị trạng thái hệ thống.
    /// </summary>
    public partial class Form1 : Form
    {
        // Giới hạn 1000 dòng nhật ký trên màn hình để tiết kiệm RAM.
        private int MaxLogLines = 1000;

        // Cơ chế khóa luồng siêu cấp (ngăn chặn các tuyến trình giẫm đạp lên nhau)
        private static SemaphoreSlim _fileLock = new SemaphoreSlim(1, 1);
        
        private readonly AppConfig _config;           // Thông số cấu hình (Thư mục, SQL).
        private readonly DatabaseService _dbService; // Dịch vụ SQL Server.
        private readonly ExcelService _excelService; // Dịch vụ đọc file Excel đo đạc.
        private FileSystemWatcher _watcher;          // "Cảm biến" cảm nhận file mới sinh.
        private bool _isProcessing;                  // Cờ ngăn việc quét tệp chồng chéo.
        private bool _isSyncingHistory = false;      // Cờ tránh việc người dùng bấm quét trùng lặp.
        private System.Windows.Forms.Timer _healthTimer; // Đồng hồ 10 giây khám sức khỏe app.
        private System.Windows.Forms.Timer _syncTimer;   // Đồng hồ quét tệp định kỳ.
        private string _lastHealthState = "";        // Ghi lại lỗi lần cuối để tránh spam log.
        private bool _isSystemHealthy = false;       // App đang ổn (True) hay đang lỗi (False).
        private NotifyIcon _trayIcon;                 // Biểu tượng nhỏ chạy dưới góc khay Windows.
        
        // Danh sách các tệp đang được "chăm sóc" (Chống trùng lặp sự kiện)
        private readonly ConcurrentDictionary<string, byte> _activeFiles = new ConcurrentDictionary<string, byte>();

        // HÀM KHỞI TẠO: Chạy đầu tiên khi bật phần mềm.
        public Form1()
        {
            InitializeComponent(); 
            
            // 1. Cài đặt chế độ tự vẽ màu cho bảng nhật ký (Matrix-Style).
            lstLogs.DrawMode = DrawMode.OwnerDrawFixed;
            lstLogs.DrawItem += LstLogs_DrawItem; 
            
            // 2. Nạp cấu hình từ tệp appsettings.json.
            _config = new AppConfig();
            _config.Load(Log);
            _config.Save(); // Tự động mã hóa mật khẩu nếu nó đang là dạng text thuần.

            // 3. Khởi động các dịch vụ phụ trợ.
            _dbService = new DatabaseService(_config, Log); 
            _excelService = new ExcelService(Log);          

            // 4. Tạo biểu tượng chạy ngầm dưới khay đồng hồ.
            _trayIcon = new NotifyIcon()
            {
                Icon = this.Icon ?? SystemIcons.Application,
                Text = "Dịch vụ Auto Import (Đang chạy ngầm)",
                Visible = true
            };

            // Nhấp đúp và icon khay để hiện lại ứng dụng.
            _trayIcon.DoubleClick += (s, e) => {
                this.Show(); 
                this.WindowState = FormWindowState.Normal; 
            };

            this.ShowInTaskbar = true; 
            this.WindowState = FormWindowState.Normal; 

            // Gán sự kiện khi hiện giao diện và khi người dùng muốn tắt app.
            this.Shown += Form1_Shown; 
            this.FormClosing += Form1_FormClosing;  
        }

        // Khi Form đã hiện lên: Bắt đầu canh gác!
        private async void Form1_Shown(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.BringToFront();
            this.Activate();

            Log($"[KHỞI TẠO] Ứng dụng đã khởi động thành công trên màn hình.");
            Log($"[KHỞI TẠO] Theo dõi thư mục: {_config.BaseFolder}");
            
            // Cài đặt đồng hồ 10 giây lặp lại.
            _healthTimer = new System.Windows.Forms.Timer { Interval = 10000 };
            _healthTimer.Tick += async (s, ev) => await PerformHealthCheckAsync(); 
            
            // Đăng ký app cùng Windows khởi động máy tính.
            RegisterAutoStart();
            
            // Khám bệnh lần 1 và bắt đầu đập nhịp đồng hồ.
            await PerformHealthCheckAsync(); 
            _healthTimer.Start();

            // Cài đặt đồng hồ quét định kỳ (mặc định 10 phút).
            _syncTimer = new System.Windows.Forms.Timer { Interval = _config.ScanIntervalSeconds * 1000 };
            _syncTimer.Tick += async (s, ev) => await SynchronizeAsync();
            _syncTimer.Start();
        }

        /// <summary>
        /// Hàm khám sức khỏe hệ thống: Chạy lặp lại mỗi 10 giây.
        /// </summary>
        private async Task PerformHealthCheckAsync()
        {
            string previousFolder = _config.BaseFolder;
            
            // Không thực hiện Load() ở đây để tránh ghi đè giá trị Folder người dùng vừa chọn bằng tay.
            // Chỉ Load cấu hình định kỳ khi hệ thống đang ở trạng thái nhàn rỗi.

            string currentState = "HEALTHY";
            bool isDirectoryOk = Directory.Exists(_config.BaseFolder); // Thư mục còn sống không?
            bool isDatabaseOk = isDirectoryOk && await _dbService.TestConnectionAsync(); // SQL còn sống không?

            if (!isDirectoryOk) currentState = "PATH_ERROR";
            else if (!isDatabaseOk) currentState = "DB_ERROR";

            _isSystemHealthy = isDirectoryOk && isDatabaseOk;

            // Xử lý thay đổi tình trạng bệnh tật của hệ thống.
            if (currentState != _lastHealthState)
            {
                string previousState = _lastHealthState;
                _lastHealthState = currentState; // Cập nhật ngay để tránh re-entry spam từ timer 10s

                if (_isSystemHealthy) 
                {
                    if (previousState != "") Log("[OK] Hệ thống đã phục hồi kết nối.");
                    UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
                    RestartWatcher(); // Kích hoạt lại cảm biến.
                    await SynchronizeAsync(); // Quét đồng bộ các tệp cũ còn sót lại.
                }
                else 
                {
                    StopWatcher(); // Hệ thống lỗi thì ngưng cảm biến cho nhẹ máy.
                    if (!isDirectoryOk) 
                    {
                        Log($"[LỖI] Không tìm thấy đường dẫn: {_config.BaseFolder}"); 
                        UpdateStatus("Lỗi Thư mục", Color.Red); 
                    }
                    else if (!isDatabaseOk) 
                    {
                        Log("[LỖI] Kết nối SQL Server thất bại.");
                        UpdateStatus("Lỗi kết nối SQL", Color.Red);
                    }
                }
            }
        }

        // Hàm khởi động lại cảm biến (Watcher) canh gác file mới.
        private void RestartWatcher()
        {
            StopWatcher();
            if (!Directory.Exists(_config.BaseFolder)) 
            {
                Directory.CreateDirectory(_config.BaseFolder); 
            }

            _watcher = new FileSystemWatcher(_config.BaseFolder) 
            {
                IncludeSubdirectories = true, // Canh gác cả thư mục con bên trong.
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite, // Nhận báo khi có file mới hoặc ghi thêm.
                Filter = "*.*",
                EnableRaisingEvents = true 
            };

            // Gắn sự kiện khi Watcher phát hiện sự thay đổi.
            _watcher.Created += OnFileEvent;
            _watcher.Changed += OnFileEvent;
            _watcher.Renamed += OnFileEvent;
        }

        // Hàm dừng cảm biến.
        private void StopWatcher()
        {
            if (_watcher != null) 
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null; 
            }
        }

        // Sự kiện khi có tiếng chuông báo động 'Có File Mới' từ Cảm biến OS.
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true; // Chặn lệnh tắt hẳn của Windows.
                this.Hide();     // Ẩn cửa sổ thôi.
                // Hiện bóng thông báo báo hiệu app vẫn sống ngầm dưới khay.
                _trayIcon.ShowBalloonTip(2000, "Import Data", "Ứng dụng vẫn đang chạy ngầm để canh file mới.", ToolTipIcon.Info); 
            }
            else
            {
                _healthTimer?.Stop();
                _syncTimer?.Stop();
                StopWatcher();
            }
        }

        // Sự kiện khi nhấn nút bấm "Thay đổi thư mục".
        private async void BtnChangeFolder_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Chọn thư mục máy đo sinh ra dữ liệu";
                dialog.SelectedPath = _config.BaseFolder;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _config.BaseFolder = dialog.SelectedPath;
                    _config.Save(); // Lưu đè vào appsettings.json
                    Log($"[CHỌN THƯ MỤC] Đã đổi đường dẫn: {_config.BaseFolder}");
                    await PerformHealthCheckAsync();
                }
            }
        }

        // Hàm cập nhật dòng chữ Trạng thái (Xanh lá, Đỏ, Vàng).
        private void UpdateStatus(string message, Color color)
        {
            // Bảo vệ đa luồng (Cross-Thread safety).
            if (this.InvokeRequired) 
            {
                this.Invoke(new Action(() => UpdateStatus(message, color))); 
                return;
            }
            lblStatus.Text = message; 
            lblStatus.ForeColor = color; 
        }

        // Hàm Tự vẽ màu cho bảng nhật ký để có hiệu ứng phát sáng.
        private void LstLogs_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            string text = lstLogs.Items[e.Index].ToString(); 
            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            // 1. VẼ NỀN (Màu đen hoặc màu Highlight xanh).
            if (isSelected)
            {
                using (var backBrush = new SolidBrush(Color.FromArgb(0, 120, 215))) 
                    e.Graphics.FillRectangle(backBrush, e.Bounds);
            }
            else
            {
                e.Graphics.FillRectangle(Brushes.Black, e.Bounds);
            }

            // 2. VẼ CHỮ.
            int timeEndIndex = text.IndexOf(']');
            string timePart = timeEndIndex > 0 ? text.Substring(0, timeEndIndex + 1) : "";
            string msgPart = timeEndIndex > 0 ? text.Substring(timeEndIndex + 1) : text;
            
            Brush textBrush = isSelected ? Brushes.White : Brushes.Lime;
            // Vẽ Thời gian và Nội dung nhật ký ở các tọa độ lệch nhau để tạo hàng lối.
            e.Graphics.DrawString(timePart, e.Font, textBrush, new PointF(e.Bounds.X + 5, e.Bounds.Y + 2));
            e.Graphics.DrawString(msgPart, e.Font, textBrush, new PointF(e.Bounds.X + 150, e.Bounds.Y + 2));
        }

        // Hàm đẩy một dòng tin nhắn vào bảng Log.
        private void Log(string message)
        {
            if (lstLogs.InvokeRequired) 
            {
                lstLogs.Invoke(new Action(() => Log(message))); 
                return;
            }
            
            string timestamp = DateTime.Now.ToString("dd/MM HH:mm:ss");
            lstLogs.Items.Add($"[{timestamp}] {message}"); 
            lstLogs.SelectedIndex = lstLogs.Items.Count - 1; 
            
            // Xóa dòng Log cũ nhất nếu vượt quá 1000 dòng.
            if (lstLogs.Items.Count > MaxLogLines) 
                lstLogs.Items.RemoveAt(0); 

            // Ghi Log ra file vật lý để đối soát.
            LogToFile($"[{timestamp}] {message}");
        }

        // Hàm ghi log ra file .txt hàng ngày.
        private void LogToFile(string logMessage)
        {
            try {
                string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
                if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
                string logFile = Path.Combine(logDir, $"log_{DateTime.Now:yyyyMMdd}.txt");
                File.AppendAllText(logFile, logMessage + Environment.NewLine);
            } catch { /* Bỏ qua nếu lỗi ghi file */ }
        }

        // Khi có tiếng chuông báo động 'Có File Mới' từ Cảm biến OS.
        private async void OnFileEvent(object sender, FileSystemEventArgs e)
        {
            // 1. Chỉ quan tâm tệp Excel (.xlsx, .xls, .xlsm) và bỏ qua các file tạm thời (bắt đầu bằng ~ hoặc $, chứa ~$).
            string ext = Path.GetExtension(e.FullPath).ToLower();
            if (ext != ".xlsx" && ext != ".xls" && ext != ".xlsm") return;

            string fileName = Path.GetFileName(e.FullPath);
            if (fileName.StartsWith("~") || fileName.Contains("~$") || fileName.StartsWith("$")) return;

            // 2. Kiểm tra trạng thái nạp file bằng Path tệp chuẩn.
            try
            {
                if (await _dbService.IsFileImportedAsync(e.FullPath)) return;
            }
            catch (Exception ex)
            {
                Log($"[CẢNH BÁO] Không kiểm tra được lịch sử file, bỏ qua: {ex.Message}");
                return;
            }

            // 3. Quan tâm tệp nằm trong thư mục ngày hôm nay (yyyy-MM-dd).
            string todayFolder = DateTime.Now.ToString("yyyy-MM-dd");
            if (!e.FullPath.Contains(todayFolder)) return; 

            // 4. Nếu file này đang được một luồng khác xử lý rồi thì bỏ qua ngay (Anti-Spam)
            if (!_activeFiles.TryAdd(e.FullPath, 0)) return;

            // Bắn vào hàng đợi tiến trình không đồng bộ.
            _ = Task.Run(async () => {
                try {
                    await ProcessSingleFileAsync(e.FullPath);
                } finally {
                    _activeFiles.TryRemove(e.FullPath, out _);
                }
            });
        }

        // Đội quân Thu Vén: Quét sạch các file còn sót trong ngày hôm nay.
        private async Task SynchronizeAsync()
        {
            if (!_isSystemHealthy || _isProcessing) return; 
            _isProcessing = true; // Chốt cửa: Không cho phép 2 đợt đồng bộ chạy song song.
            // Quét dữ liệu của 2 ngày: Hôm nay và Hôm qua để tránh sót dữ liệu khi qua đêm.
            string today = DateTime.Now.ToString("yyyy-MM-dd");
            string yesterday = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            string[] targetFolders = { yesterday, today };

            UpdateStatus("Đang quét tệp...", Color.Yellow); 

            try
            {
                var filesToProcess = new List<string>();
                foreach (var folder in targetFolders)
                {
                    string sourcePath = Path.Combine(_config.BaseFolder, folder);
                    if (!Directory.Exists(sourcePath)) continue;

                    string[] found = await Task.Run(() => 
                        Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories)
                            .Where(f => f.EndsWith(".xlsx") || f.EndsWith(".xls") || f.EndsWith(".xlsm"))
                            .Where(f => {
                                string name = Path.GetFileName(f);
                                return !name.StartsWith("~") && !name.Contains("~$") && !name.StartsWith("$");
                            })
                            .ToArray()
                    );
                    filesToProcess.AddRange(found);
                }

                if (filesToProcess.Count == 0)
                {
                    UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
                    return;
                }

                // Ném TOÀN BỘ vòng lặp xuống Background Worker.
                await Task.Run(async () => 
                {
                    int count = 0;
                    for (int i = 0; i < filesToProcess.Count; i++)
                    {
                        string file = filesToProcess[i];
                        UpdateStatus($"Đồng bộ {++count}/{filesToProcess.Count}", Color.Orange);
                        
                        if (_activeFiles.TryAdd(file, 0))
                        {
                            try {
                                await ProcessSingleFileAsync(file, silentIfImported: true);
                            } finally {
                                _activeFiles.TryRemove(file, out _);
                            }
                        }

                        // TỐI ƯU: Thêm khoảng nghỉ 300ms sau mỗi 10 file để SQL Server không bị quá tải CPU/IO
                        if ((i + 1) % 10 == 0 && i + 1 < filesToProcess.Count)
                        {
                            await Task.Delay(300);
                        }
                    }
                });
                
                UpdateStatus("Hệ thống Sẵn sàng", Color.Green); 
            }
            catch (Exception ex)
            {
                Log($"[LỖI] Đồng bộ hóa thất bại: {ex}"); 
                UpdateStatus("Lỗi Đồng bộ", Color.Red); 
            }
            finally
            {
                _isProcessing = false; // Mở cửa cho lần đồng bộ tiếp theo.
            }
        }

        /// <summary>
        /// TRUNG TÂM XỬ LÝ: 'Giải cứu' tệp Excel từ ổ cứng để đưa vào SQL.
        /// </summary>
        private async Task ProcessSingleFileAsync(string filePath, bool silentIfImported = false)
        {
            string fileName = Path.GetFileName(filePath);

            // Xin chìa khóa để vào phòng xử lý (Tuyệt đối không cho 2 tệp xử lý cùng 1 mili-giây).
            await _fileLock.WaitAsync();
            try
            {
                // 1. Kiểm tra trạng thái nạp file bằng Path tệp chuẩn.
                if (await _dbService.IsFileImportedAsync(filePath)) 
                {
                    if (!silentIfImported) Log($"[BỎ QUA] Tệp đã được nạp từ trước: {fileName}");
                    return;
                }

                // 2. Đàm phán O.S: Chờ máy đo buông tay khỏi tệp (Ready to Read).
                if (!await IsFileReadyAsync(filePath))
                {
                    // Chỉ Log "Bận" khi không phải quét tự động, giúp màn hình sạch sẽ hơn.
                    if (!silentIfImported) Log($"[TẠM DỪNG] Tệp đang bị máy đo giữ: {fileName}");
                    return;
                }

                Log($"[ĐANG NẠP] Xử lý tệp: {fileName}");

                // 3. Đọc dữ liệu Excel bọc trong Task.Run để giải phóng Thread UI (chống Not Responding).
                var data = await Task.Run(() => _excelService.ReadExcelFile(filePath));
                if (data == null || data.Rows.Count == 0) return; 

                // 4. Nhập hàng loạt vào SQL.
                await _dbService.ExecuteImportBatchAsync(data, fileName, filePath);
            }
            catch (Exception ex)
            {
                Log($"[LỖI] Xử lý {fileName} bị dừng giữa chừng: {ex.Message}"); 
            }
            finally
            {
                _fileLock.Release(); // Xử lý xong thì trả chìa khóa cho thằng tiếp theo.
            }
        }

        // Hàm giúp app chờ đợi file khi nó đang bị máy đo 'rặn' nốt dữ liệu.
        internal async Task<bool> IsFileReadyAsync(string filePath)
        {
            // Thử 10 lần với khoảng nghỉ 3s (tổng cộng 30 giây) để chờ máy đo hoặc công nhân nhả file.
            // Đây là khoảng thời gian "vàng" đảm bảo cân bằng giữa hiệu suất và độ tin cậy.
            for (int i = 0; i < 10; i++)
            {
                if (!File.Exists(filePath)) return false; // Nếu file bị xóa/di chuyển mất thì dừng ngay lập tức
                try
                {
                    // Thử mượn file với quyền độc quyền (None).
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                        return true; // Thành công - file đã sẵn sàng.
                }
                catch (IOException) 
                {
                    // Chờ thêm 3 giây rồi gõ cửa lại lượt tiếp theo.
                    await Task.Delay(3000); 
                }
            }
            return false;
        }

        // Tự đăng ký app vào thư mục Startup của Registry Windows.
        private void RegisterAutoStart()
        {
            try 
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key != null) key.SetValue("AutoImportData", Application.ExecutablePath);
                }
            }
            catch (Exception ex) 
            {
                Log($"[CẢNH BÁO] Không thể đặt tự khởi động: {ex.Message}");
            }
        }

        // Sự kiện khi nhấn nút bấm "Đồng bộ lịch sử".
        private async void BtnSyncHistory_Click(object sender, EventArgs e)
        {
            if (!_isSystemHealthy)
            {
                Log("[CẢNH BÁO] Hệ thống đang có lỗi (SQL hoặc thư mục), vui lòng kiểm tra trước khi đồng bộ.");
                MessageBox.Show("Hệ thống đang có lỗi. Vui lòng kiểm tra kết nối SQL hoặc đường dẫn thư mục!", "Lỗi kết nối", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_isSyncingHistory)
            {
                Log("[CẢNH BÁO] Tiến trình quét lịch sử đang chạy ngầm, không bấm liên tục.");
                return;
            }

            using (var dialog = new SyncHistoryDialog())
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var options = dialog.Options;
                    await Task.Run(async () => await SyncAllHistoryAsync(options));
                }
            }
        }



        /// <summary>
        /// Tiến trình quét và đồng bộ dữ liệu lịch sử ngầm.
        /// TỐI ƯU: Load toàn bộ lịch sử import 1 lần vào RAM (1 query), lọc file cần nạp local.
        /// Có cơ chế batch processing và circuit breaker.
        /// </summary>
        private async Task SyncAllHistoryAsync(SyncOptions options)
        {
            _isSyncingHistory = true;
            UpdateStatus("Đang quét lịch sử...", Color.Yellow);
            
            string modeText = options.Mode switch
            {
                SyncMode.LastNDays => $"gần đây ({options.Days} ngày)",
                SyncMode.DateRange => $"khoảng từ {options.StartDate:dd/MM/yyyy} đến {options.EndDate:dd/MM/yyyy}",
                _ => "toàn bộ lịch sử"
            };
            Log($"[ĐỒNG BỘ LỊCH SỬ] Bắt đầu đồng bộ {modeText} dưới thư mục gốc...");

            try
            {
                string baseDir = _config.BaseFolder;
                if (!Directory.Exists(baseDir))
                {
                    Log($"[LỖI] Không tìm thấy thư mục gốc: {baseDir}");
                    UpdateStatus("Lỗi Thư mục", Color.Red);
                    return;
                }

                // 1. Quét tìm các thư mục có định dạng ngày yyyy-MM-dd
                var regex = new System.Text.RegularExpressions.Regex(@"^\d{4}-\d{2}-\d{2}$");
                var subDirs = Directory.GetDirectories(baseDir)
                    .Select(Path.GetFileName)
                    .Where(name => name != null && regex.IsMatch(name))
                    .OrderByDescending(name => name) // Quét từ ngày mới nhất ngược về trước
                    .ToList();

                // Lọc danh sách thư mục theo cấu hình đồng bộ
                if (options.Mode == SyncMode.LastNDays)
                {
                    DateTime limitDate = DateTime.Today.AddDays(-options.Days);
                    subDirs = subDirs.Where(name =>
                    {
                        if (DateTime.TryParseExact(name, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime folderDate))
                        {
                            return folderDate >= limitDate;
                        }
                        return false;
                    }).ToList();
                }
                else if (options.Mode == SyncMode.DateRange)
                {
                    subDirs = subDirs.Where(name =>
                    {
                        if (DateTime.TryParseExact(name, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime folderDate))
                        {
                            return folderDate >= options.StartDate.Date && folderDate <= options.EndDate.Date;
                        }
                        return false;
                    }).ToList();
                }

                if (subDirs.Count == 0)
                {
                    Log("[ĐỒNG BỘ LỊCH SỬ] Không tìm thấy thư mục ngày yyyy-MM-dd nào phù hợp với phạm vi lọc.");
                    UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
                    return;
                }

                Log($"[ĐỒNG BỘ LỊCH SỬ] Tìm thấy {subDirs.Count} thư mục ngày đo để kiểm tra.");

                // 2. Tìm tất cả tệp Excel trên ổ cứng trong các thư mục được lọc
                UpdateStatus("Đang quét tệp trên ổ cứng...", Color.Yellow);
                var allFiles = new List<string>();
                foreach (var dirName in subDirs)
                {
                    string fullPath = Path.Combine(baseDir, dirName);
                    if (Directory.Exists(fullPath))
                    {
                        var files = Directory.GetFiles(fullPath, "*.*", SearchOption.AllDirectories)
                            .Where(f => f.EndsWith(".xlsx") || f.EndsWith(".xls") || f.EndsWith(".xlsm"))
                            .Where(f => {
                                string name = Path.GetFileName(f);
                                return !name.StartsWith("~") && !name.Contains("~$") && !name.StartsWith("$");
                            })
                            .ToArray();
                        allFiles.AddRange(files);
                    }
                }

                if (allFiles.Count == 0)
                {
                    Log("[ĐỒNG BỘ LỊCH SỬ] Không tìm thấy tệp Excel nào trong các thư mục lịch sử được lọc.");
                    UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
                    return;
                }

                Log($"[ĐỒNG BỘ LỊCH SỬ] Tìm thấy {allFiles.Count} tệp Excel trên ổ cứng.");

                var filesToProcess = new List<string>();
                if (options.OverwriteExisting)
                {
                    filesToProcess = allFiles;

                    // Thực hiện xóa dữ liệu cũ của tất cả file đo được chọn để chuẩn bị nạp lại
                    UpdateStatus("Đang xóa dữ liệu cũ...", Color.OrangeRed);
                    Log($"[ĐỒNG BỘ LỊCH SỬ] Đang xóa dữ liệu cũ của {filesToProcess.Count} tệp trong database...");
                    var (deletedRows, deletedHistory) = await _dbService.DeleteByFilePathsAsync(filesToProcess);
                    Log($"[ĐỒNG BỘ LỊCH SỬ] Đã xóa {deletedRows:N0} dòng dữ liệu và {deletedHistory:N0} bản ghi lịch sử.");
                }
                else
                {
                    // Lọc và chỉ nạp các file chưa được nạp thành công để tiết kiệm thời gian
                    UpdateStatus("Đang đối soát tệp đã nạp...", Color.Yellow);
                    var (importedPaths, fileNameSizeMap) = await _dbService.GetAllImportedFilesAsync();
                    
                    foreach (var file in allFiles)
                    {
                        bool isAlreadyImported = false;
                        if (importedPaths.Contains(file))
                        {
                            isAlreadyImported = true;
                        }
                        else
                        {
                            // Kiểm tra theo tên file + dung lượng
                            string name = Path.GetFileName(file);
                            if (fileNameSizeMap.TryGetValue(name, out long size))
                            {
                                try
                                {
                                    long currentSize = new FileInfo(file).Length;
                                    if (currentSize == size)
                                    {
                                        isAlreadyImported = true;
                                    }
                                }
                                catch { }
                            }
                        }

                        if (!isAlreadyImported)
                        {
                            filesToProcess.Add(file);
                        }
                    }

                    if (filesToProcess.Count == 0)
                    {
                        Log("[ĐỒNG BỘ LỊCH SỬ] Tất cả các tệp trong phạm vi lọc đều đã được nạp trước đó. Không có tệp mới cần đồng bộ.");
                        UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
                        return;
                    }

                    Log($"[ĐỒNG BỘ LỊCH SỬ] Phát hiện {filesToProcess.Count}/{allFiles.Count} tệp chưa được nạp (Bỏ qua {allFiles.Count - filesToProcess.Count} tệp đã có).");
                }

                int importedSuccessCount = 0;
                int consecutiveErrors = 0;
                const int BatchSize = 10;
                const int BatchDelayMs = 500;
                const int CircuitBreakerThreshold = 3;
                const int CircuitBreakerDelayMs = 30000;

                for (int i = 0; i < filesToProcess.Count; i++)
                {
                    string file = filesToProcess[i];
                    string fileName = Path.GetFileName(file);
                    UpdateStatus($"Nạp lại {i + 1}/{filesToProcess.Count}", Color.Orange);

                    try
                    {
                        Log($"[ĐỒNG BỘ LỊCH SỬ] Nạp tệp: {fileName} ({Path.GetFileName(Path.GetDirectoryName(file))})");

                        if (_activeFiles.TryAdd(file, 0))
                        {
                            try
                            {
                                await ProcessSingleFileAsync(file, silentIfImported: false);
                                importedSuccessCount++;
                                consecutiveErrors = 0;
                            }
                            finally
                            {
                                _activeFiles.TryRemove(file, out _);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        consecutiveErrors++;
                        Log($"[LỖI] Lỗi nạp tệp {fileName}: {ex.Message}");

                        if (consecutiveErrors >= CircuitBreakerThreshold)
                        {
                            Log($"[CIRCUIT BREAKER] Đã gặp {consecutiveErrors} lỗi liên tiếp. Tạm dừng {CircuitBreakerDelayMs / 1000} giây...");
                            UpdateStatus("Chờ SQL phục hồi...", Color.Yellow);
                            await Task.Delay(CircuitBreakerDelayMs);
                            consecutiveErrors = 0;
                            Log("[CIRCUIT BREAKER] Tiếp tục đồng bộ...");
                        }
                    }

                    // Batch delay: Mỗi 10 tệp, nghỉ 500ms
                    if ((i + 1) % BatchSize == 0 && i + 1 < filesToProcess.Count)
                    {
                        await Task.Delay(BatchDelayMs);
                    }
                }

                Log($"[ĐỒNG BỘ LỊCH SỬ] Hoàn tất quá trình đồng bộ!");
                Log($"[ĐỒNG BỘ LỊCH SỬ] Kết quả: {filesToProcess.Count} tệp tổng, nạp thành công {importedSuccessCount} tệp.");
                UpdateStatus("Hệ thống Sẵn sàng", Color.Green);
            }
            catch (Exception ex)
            {
                Log($"[LỖI] Đồng bộ lịch sử thất bại: {ex.Message}");
                UpdateStatus("Lỗi Đồng bộ", Color.Red);
            }
            finally
            {
                _isSyncingHistory = false;
            }
        }
    }
}
