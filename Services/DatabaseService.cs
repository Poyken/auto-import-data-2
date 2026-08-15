using Microsoft.Data.SqlClient;
using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using ImportData.Core;

namespace ImportData.Services 
{
    public class DatabaseService
    {
        private const string TableData = "SortingDataImportExcel";
        private const string TableHistory = "ExcelImportHistory";

        // Số lần retry tối đa khi gặp lỗi SQL transient
        private const int MaxRetryCount = 3;
        // Delay cơ bản giữa các lần retry (sẽ nhân lên theo cấp số nhân)
        private static readonly int[] RetryDelaysMs = { 2000, 5000, 10000 };

        private static readonly string[] SqlColumns = {
            "EquipmentNumber", "SorterNum", "StartTime", "WorkflowCode",
            "Barcode", "Slot", "Position", "Channel", "Capacity_mAh", "Capacitance_F", 
            "BeginVoltageSD_mV", "ChargeEndCurrent_mA", "EndVoltage_mV", "EndCurrent_mA", "DischargeVoltage1_mV", 
            "DischargeVoltage1_Time", "DischargeVoltage2_mV", "DischargeVoltage2_Time", "DischargeBeginVoltage_mV", "DischargeBeginCurrent_mA", 
            "NGInfo", "EndTime", "FilePath", "ImportDate", "ESR_mOhm", "OCV_mV", "ESRTime"
        };

        /// <summary>
        /// Chuẩn hóa tên cột: Loại bỏ ký tự điều khiển, Unicode ẩn, khoảng trắng, dấu đặc biệt.
        /// Dùng để so khớp tên cột Excel với alias SQL bất kể định dạng máy đo nào.
        /// </summary>
        internal static string GetSearchKey(string name)
        {
            if (string.IsNullOrEmpty(name)) return string.Empty;
            // Loại bỏ ký tự điều khiển (newline, tab...), Zero-Width Space (\u200B), BOM (\uFEFF)
            var chars = name.Where(c => !char.IsControl(c) && (int)c != 8203 && (int)c != 65279).ToArray();
            string clean = new string(chars).Trim();
            return clean.Replace(" ", "").Replace("_", "").Replace("-", "")
                        .Replace("(", "").Replace(")", "").Replace("/", "").Replace("\\", "")
                        .Replace("%", "").Replace(":", "").Replace("：", "").Replace(";", "").Replace(",", "").ToLower();
        }

        /// <summary>
        /// Bảng alias: Ánh xạ từ tên cột Excel (đã chuẩn hóa) sang tên cột SQL.
        /// Hỗ trợ cả 2 định dạng máy đo: Format 1 (Equipment Number) và Format 2 (DevName).
        /// </summary>
        internal static readonly Dictionary<string, string> AliasToSqlColumnMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // --- Cột thiết bị (Equipment) ---
            { "equipmentnumber", "EquipmentNumber" },
            { "devname", "EquipmentNumber" },
            { "eqpno", "EquipmentNumber" },
            { "machineid", "EquipmentNumber" },
            // --- Cột Sorter ---
            { "sorternum", "SorterNum" },
            { "sortnum", "SorterNum" },
            // --- Các cột dữ liệu đo ---
            { "starttime", "StartTime" },
            { "workflowcode", "WorkflowCode" },
            { "barcode", "Barcode" },
            { "lotno", "Barcode" },
            { "lotid", "Barcode" },
            { "cellid", "Barcode" },
            { "slot", "Slot" },
            { "position", "Position" },
            { "channel", "Channel" },
            { "capacitymah", "Capacity_mAh" },
            { "capmah", "Capacity_mAh" },
            { "capacity", "Capacity_mAh" },
            { "capacitancef", "Capacitance_F" },
            { "capf", "Capacitance_F" },
            { "capacitance", "Capacitance_F" },
            { "beginvoltagesdmv", "BeginVoltageSD_mV" },
            { "beginvoltagesd", "BeginVoltageSD_mV" },
            { "chargeendcurrentma", "ChargeEndCurrent_mA" },
            { "chargeendcurrent", "ChargeEndCurrent_mA" },
            { "endvoltagemv", "EndVoltage_mV" },
            { "endvoltage", "EndVoltage_mV" },
            { "endcurrentma", "EndCurrent_mA" },
            { "endcurrent", "EndCurrent_mA" },
            { "dischargevoltage1mv", "DischargeVoltage1_mV" },
            { "dischargevoltage1", "DischargeVoltage1_mV" },
            { "dischargevoltage2mv", "DischargeVoltage2_mV" },
            { "dischargevoltage2", "DischargeVoltage2_mV" },
            { "dischargebeginvoltagemv", "DischargeBeginVoltage_mV" },
            { "dischargebeginvoltage", "DischargeBeginVoltage_mV" },
            { "dischargebegincurrentma", "DischargeBeginCurrent_mA" },
            { "dischargebegincurrent", "DischargeBeginCurrent_mA" },
            { "esrmω", "ESR_mOhm" },
            { "esrmohm", "ESR_mOhm" },
            { "esr", "ESR_mOhm" },
            { "ocvmv", "OCV_mV" },
            { "ocv", "OCV_mV" },
            { "esrtime", "ESRTime" },
            { "nginfo", "NGInfo" },
            { "endtime", "EndTime" }
        };

        private readonly AppConfig _config;
        private readonly Action<string> _logger;
        private string _lastConnectionString;

        public DatabaseService(AppConfig config, Action<string> logger)
        {
            _config = config; 
            _logger = logger; 
            _lastConnectionString = config.ConnectionString; 
        }

        /// <summary>
        /// Tạo connection string với các tham số pool phù hợp cho batch processing.
        /// </summary>
        private string GetOptimizedConnectionString(int timeout = 0)
        {
            var builder = new SqlConnectionStringBuilder(_config.ConnectionString)
            {
                ConnectTimeout = timeout > 0 ? timeout : 15,
                MaxPoolSize = 20,
                MinPoolSize = 1,
                Pooling = true
            };
            return builder.ConnectionString;
        }

        /// <summary>
        /// Tạo kết nối SQL mới với cơ chế retry tự động.
        /// Khi mất kết nối, sẽ thử lại tối đa 3 lần với delay tăng dần (2s, 5s, 10s).
        /// </summary>
        private async Task<SqlConnection> CreateConnectionWithRetryAsync()
        {
            // Phát hiện thay đổi connection string → xóa pool cũ
            if (_config.ConnectionString != _lastConnectionString)
            {
                SqlConnection.ClearAllPools();
                _lastConnectionString = _config.ConnectionString;
            }

            for (int attempt = 0; attempt < MaxRetryCount; attempt++)
            {
                SqlConnection conn = null;
                try
                {
                    conn = new SqlConnection(GetOptimizedConnectionString());
                    await conn.OpenAsync();
                    return conn;
                }
                catch (SqlException ex) when (IsTransientError(ex) && attempt < MaxRetryCount - 1)
                {
                    conn?.Dispose();
                    int delay = RetryDelaysMs[attempt];
                    _logger?.Invoke($"[SQL-RETRY] Lỗi kết nối (lần {attempt + 1}/{MaxRetryCount}), chờ {delay / 1000}s... ({ex.Message})");
                    await Task.Delay(delay);
                }
                catch
                {
                    conn?.Dispose();
                    throw;
                }
            }
            // Lần cuối cùng: Nếu vẫn lỗi thì throw ra ngoài
            var finalConn = new SqlConnection(GetOptimizedConnectionString());
            await finalConn.OpenAsync();
            return finalConn;
        }

        /// <summary>
        /// Kiểm tra xem mã lỗi SQL có phải lỗi tạm thời (mạng, timeout) có thể retry được không.
        /// </summary>
        internal static bool IsTransientErrorNumber(int errorNumber)
        {
            // Các mã lỗi SQL transient phổ biến
            int[] transientErrors = { -2, 20, 64, 233, 10053, 10054, 10060, 40143, 40197, 40501, 40613, 49918, 49919, 49920 };
            return transientErrors.Contains(errorNumber) || errorNumber == -1 || errorNumber == 258;
        }

        private static bool IsTransientError(SqlException ex)
        {
            return IsTransientErrorNumber(ex.Number);
        }

        public async Task<bool> TestConnectionAsync()
        {
            try 
            {
                if (_config.ConnectionString != _lastConnectionString)
                {
                    SqlConnection.ClearAllPools();
                    _lastConnectionString = _config.ConnectionString;
                }
                var builder = new SqlConnectionStringBuilder(_config.ConnectionString) { ConnectTimeout = _config.HealthCheckTimeoutSeconds };
                using (var testConn = new SqlConnection(builder.ConnectionString)) 
                {
                    await testConn.OpenAsync();
                    return true;
                } 
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI-SQL-CONNECT] {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Nạp TOÀN BỘ lịch sử import thành công từ SQL vào bộ nhớ RAM trong 1 query duy nhất.
        /// Trả về HashSet chứa đường dẫn file đã import + Dictionary (tên file → dung lượng) để kiểm tra file bị di chuyển.
        /// Dùng cho Sync History để tránh gọi 10,000 query SQL riêng lẻ.
        /// </summary>
        public async Task<(HashSet<string> importedPaths, Dictionary<string, long> fileNameSizeMap)> GetAllImportedFilesAsync()
        {
            var importedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var fileNameSizeMap = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (var conn = await CreateConnectionWithRetryAsync())
                {
                    string sql = $"SELECT FilePath, FileSize FROM {TableHistory} WHERE Status = 'Success'";
                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 120;
                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            while (await reader.ReadAsync())
                            {
                                string path = reader.GetString(0);
                                long size = reader.IsDBNull(1) ? 0 : reader.GetInt64(1);
                                
                                importedPaths.Add(path);
                                
                                // Lưu tên file + dung lượng để kiểm tra file bị di chuyển folder
                                string fileName = Path.GetFileName(path);
                                if (!fileNameSizeMap.ContainsKey(fileName))
                                {
                                    fileNameSizeMap[fileName] = size;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI-SQL] Không tải được lịch sử import: {ex.Message}");
                throw;
            }

            return (importedPaths, fileNameSizeMap);
        }

        /// <summary>
        /// Kiểm tra xem tệp Excel này đã được nạp thành công vào hệ thống trước đó chưa.
        /// Chống nạp trùng thông minh: Kiểm tra theo Đường dẫn HOẶC (Tên file + Dung lượng).
        /// Có retry logic để tránh crash khi mạng không ổn định.
        /// </summary>
        public async Task<bool> IsFileImportedAsync(string filePath) 
        {
            try
            {
                if (!File.Exists(filePath)) return false;
                
                long fileSize = new FileInfo(filePath).Length;

                // Trích xuất đường dẫn tương đối từ BaseFolder để tránh trùng lặp chéo giữa các thiết bị (ví dụ 11# và 12#)
                string relativePath = filePath;
                if (!string.IsNullOrEmpty(_config.BaseFolder) && filePath.StartsWith(_config.BaseFolder, StringComparison.OrdinalIgnoreCase))
                {
                    relativePath = filePath.Substring(_config.BaseFolder.Length).TrimStart('\\', '/');
                }
                // Chuẩn hóa ký tự phân tách thư mục về dạng Windows để đối chiếu trong SQL
                relativePath = relativePath.Replace('/', '\\');

                using (var conn = await CreateConnectionWithRetryAsync()) 
                {
                    // Kiểm tra 2 điều kiện:
                    // 1. Đúng đường dẫn tuyệt đối (đã nạp rồi)
                    // 2. Hoặc Cùng đường dẫn tương đối (bao gồm cả thư mục ngày và thư mục máy đo như 12#\data.xlsx) và cùng dung lượng
                    string sql = $@"SELECT COUNT(*) FROM {TableHistory} 
                                   WHERE Status = 'Success' 
                                   AND (FilePath = @path OR (FilePath LIKE '%' + @relativePath AND FileSize = @size))";
                    
                    using (var cmd = new SqlCommand(sql, conn)) 
                    {
                        cmd.CommandTimeout = 30;
                        cmd.Parameters.AddWithValue("@path", filePath);
                        cmd.Parameters.AddWithValue("@relativePath", relativePath);
                        cmd.Parameters.AddWithValue("@size", fileSize);
                        int count = Convert.ToInt32(await cmd.ExecuteScalarAsync()); 
                        return count > 0;
                    }
                } 
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI-SQL] Không kiểm tra được lịch sử nạp file: {ex.Message}");
                throw; // Ném lỗi để Form1 biết mà dừng lại, tránh nạp trùng dữ liệu khi mạng lỗi.
            }
        }

        public async Task<int> ExecuteImportBatchAsync(DataTable dt, string fileName, string filePath) 
        {
            if (dt == null || dt.Rows.Count == 0) return 0;

            using (var conn = await CreateConnectionWithRetryAsync()) 
            {
                using (var trans = conn.BeginTransaction())
                {
                    try
                    {
                        // 1. Chuẩn bị dữ liệu bổ sung (FilePath, ImportDate)
                        // ImportDate = thời gian file được máy đo ghi hoàn chỉnh (LastWriteTime), KHÔNG phải thời gian nạp.
                        if (!dt.Columns.Contains("FilePath")) dt.Columns.Add("FilePath", typeof(string));
                        if (!dt.Columns.Contains("ImportDate")) dt.Columns.Add("ImportDate", typeof(DateTime));
                        
                        DateTime fileWriteTime = DateTime.Now;
                        bool parsedFromFileName = false;
                        try
                        {
                            string nameNoExt = Path.GetFileNameWithoutExtension(filePath);
                            var match = System.Text.RegularExpressions.Regex.Match(nameNoExt, @"\d{14}");
                            if (match.Success)
                            {
                                if (DateTime.TryParseExact(match.Value, "yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime parsedTime))
                                {
                                    fileWriteTime = parsedTime;
                                    parsedFromFileName = true;
                                }
                            }
                        }
                        catch
                        {
                            // Bỏ qua nếu có lỗi parse tên file
                        }

                        if (!parsedFromFileName)
                        {
                            try
                            {
                                fileWriteTime = File.GetLastWriteTime(filePath);
                            }
                            catch
                            {
                                // Nếu không lấy được thời gian file, dùng thời gian hiện tại làm fallback
                            }
                        }

                        // Định vị các cột khoá chính để lọc dữ liệu hợp lệ
                        int barcodeColIdx = -1;
                        int startTimeColIdx = -1;
                        int equipmentColIdx = -1;
                        
                        // Tìm tất cả các cột có thể làm Barcode (ví dụ: Barcode, LotNo, LotID, CellID)
                        var barcodeColCandidates = new List<int>();
                        for (int col = 0; col < dt.Columns.Count; col++)
                        {
                            string searchKey = GetSearchKey(dt.Columns[col].ColumnName);
                            if (AliasToSqlColumnMap.TryGetValue(searchKey, out string? sqlCol))
                            {
                                if (sqlCol == "Barcode") barcodeColCandidates.Add(col);
                                else if (sqlCol == "StartTime") startTimeColIdx = col;
                                else if (sqlCol == "EquipmentNumber") equipmentColIdx = col;
                            }
                        }

                        // Chọn cột Barcode tốt nhất có chứa dữ liệu thực tế
                        if (barcodeColCandidates.Count > 0)
                        {
                            barcodeColIdx = barcodeColCandidates[0]; // Mặc định chọn cột ứng viên đầu tiên
                            
                            foreach (int colIdx in barcodeColCandidates)
                            {
                                bool hasData = false;
                                int checkLimit = Math.Min(dt.Rows.Count, 10);
                                for (int r = 0; r < checkLimit; r++)
                                {
                                    var val = dt.Rows[r][colIdx];
                                    if (val != null && !string.IsNullOrWhiteSpace(val.ToString()) && val.ToString() != "---")
                                    {
                                        hasData = true;
                                        break;
                                    }
                                }
                                
                                if (hasData)
                                {
                                    // Ưu tiên cột có chứa dữ liệu thực tế. Nếu cột có tên chính xác là "barcode" và có dữ liệu thì càng tốt.
                                    string colName = GetSearchKey(dt.Columns[colIdx].ColumnName);
                                    if (colName == "barcode" || barcodeColIdx == barcodeColCandidates[0] || string.IsNullOrWhiteSpace(dt.Rows[0][barcodeColIdx]?.ToString()))
                                    {
                                        barcodeColIdx = colIdx;
                                    }
                                }
                            }
                        }

                        // Duyệt ngược để xóa các dòng lỗi/thiếu khoá chính, gán thông tin file cho dòng hợp lệ
                        for (int i = dt.Rows.Count - 1; i >= 0; i--)
                        {
                            DataRow row = dt.Rows[i];
                            bool isValid = true;
                            
                            if (barcodeColIdx >= 0 && (row[barcodeColIdx] == DBNull.Value || string.IsNullOrWhiteSpace(row[barcodeColIdx].ToString()))) isValid = false;
                            if (startTimeColIdx >= 0 && (row[startTimeColIdx] == DBNull.Value || string.IsNullOrWhiteSpace(row[startTimeColIdx].ToString()))) isValid = false;
                            if (equipmentColIdx >= 0 && (row[equipmentColIdx] == DBNull.Value || string.IsNullOrWhiteSpace(row[equipmentColIdx].ToString()))) isValid = false;
                            
                            if (!isValid)
                            {
                                dt.Rows.RemoveAt(i);
                                continue;
                            }

                            row["FilePath"] = filePath; 
                            row["ImportDate"] = fileWriteTime; 
                        }

                        // 2. TẠO BẢNG TẠM TRONG SQL (Dùng cấu trúc tương đương bảng đích)
                        string tempTableName = $"#TempImport_{Guid.NewGuid().ToString("N")}";
                        string createTempSql = $"SELECT TOP 0 * INTO {tempTableName} FROM {TableData}";
                        using (var cmdCreate = new SqlCommand(createTempSql, conn, trans)) { await cmdCreate.ExecuteNonQueryAsync(); }

                        // 3. BULK COPY VÀO BẢNG TẠM (Tốc độ cực nhanh)
                        using (var bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trans))
                        {
                            bulkCopy.DestinationTableName = tempTableName;
                            bulkCopy.BatchSize = 1000;
                            bulkCopy.BulkCopyTimeout = 120;

                            // Thực hiện ánh xạ cột chính xác
                            for (int col = 0; col < dt.Columns.Count; col++)
                            {
                                DataColumn dc = dt.Columns[col];
                                string searchKey = GetSearchKey(dc.ColumnName);

                                // Xử lý động cho cột thời gian xả (Discharge Time) do có nhiều định dạng đơn vị khác nhau
                                if (searchKey.StartsWith("dischargevoltage1time"))
                                {
                                    bulkCopy.ColumnMappings.Add(dc.ColumnName, "DischargeVoltage1_Time");
                                    continue;
                                }
                                if (searchKey.StartsWith("dischargevoltage2time"))
                                {
                                    bulkCopy.ColumnMappings.Add(dc.ColumnName, "DischargeVoltage2_Time");
                                    continue;
                                }

                                if (AliasToSqlColumnMap.TryGetValue(searchKey, out string? sqlCol))
                                {
                                    if (sqlCol == "Barcode")
                                    {
                                        // Chỉ ánh xạ cột được chọn làm Barcode thực tế chứa dữ liệu
                                        if (col == barcodeColIdx)
                                        {
                                            bulkCopy.ColumnMappings.Add(dc.ColumnName, "Barcode");
                                        }
                                        continue;
                                    }

                                    bulkCopy.ColumnMappings.Add(dc.ColumnName, sqlCol);
                                }
                            }
                            // Bắt buộc map 2 cột metadata
                            bulkCopy.ColumnMappings.Add("FilePath", "FilePath");
                            bulkCopy.ColumnMappings.Add("ImportDate", "ImportDate");
                            
                            await bulkCopy.WriteToServerAsync(dt);
                        }

                        // 4. MERGE DỮ LIỆU TỪ BẢNG TẠM SANG BẢNG CHÍNH
                        // KEY UNIQUE đúng cho dữ liệu cell: Barcode + StartTime + EquipmentNumber + Channel + Position
                        // Lý do PHẢI có Channel + Position:
                        //   - Mỗi file Excel có 64 rows, TẤT CẢ đều cùng Barcode + StartTime + EquipmentNumber
                        //   - Nếu chỉ partition theo 3 cột trên → ROW_NUMBER() group tất cả 64 rows vào 1 nhóm
                        //     → WHERE rn=1 chỉ giữ lại 1 row duy nhất, 63 rows còn lại bị bỏ! (BUG TRƯỚC ĐÂY)
                        //   - Channel + Position phân biệt từng cell riêng lẻ trong cùng 1 batch
                        string mergeSql = $@"
                            WITH UniqueSrc AS (
                                SELECT *, ROW_NUMBER() OVER (
                                    PARTITION BY Barcode, StartTime, EquipmentNumber, Channel, Position
                                    ORDER BY (SELECT NULL)
                                ) as rn
                                FROM {tempTableName}
                            )
                            INSERT INTO {TableData} ({string.Join(", ", SqlColumns)})
                            SELECT {string.Join(", ", SqlColumns)}
                            FROM UniqueSrc AS src
                            WHERE rn = 1
                            AND NOT EXISTS (
                                SELECT 1 FROM {TableData} AS dest
                                WHERE dest.Barcode = src.Barcode 
                                AND dest.StartTime = src.StartTime
                                AND dest.EquipmentNumber = src.EquipmentNumber
                                AND dest.Channel = src.Channel
                                AND dest.Position = src.Position
                            )";
                        
                        int rowsInserted = 0;
                        using (var cmdMerge = new SqlCommand(mergeSql, conn, trans))
                        {
                            cmdMerge.CommandTimeout = 300;
                            rowsInserted = await cmdMerge.ExecuteNonQueryAsync();
                        }

                        // 5. Ghi lịch sử nạp file (Sử dụng UPSERT để tránh vi phạm khóa unique UX_ExcelImportHistory_FilePath)
                        long fileSize = new FileInfo(filePath).Length;
                        string historySql = $@"
                            IF EXISTS (SELECT 1 FROM {TableHistory} WHERE FilePath = @path)
                            BEGIN
                                UPDATE {TableHistory} 
                                SET FileSize = @size, ImportedAt = GETDATE(), RowsInserted = @rows, Status = 'Success'
                                WHERE FilePath = @path
                            END
                            ELSE
                            BEGIN
                                INSERT INTO {TableHistory} (FilePath, FileSize, ImportedAt, RowsInserted, Status) 
                                VALUES (@path, @size, GETDATE(), @rows, 'Success')
                            END";
                        
                        using (var cmd = new SqlCommand(historySql, conn, trans)) 
                        {
                            cmd.Parameters.AddWithValue("@path", filePath);   
                            cmd.Parameters.AddWithValue("@size", fileSize);   
                            cmd.Parameters.AddWithValue("@rows", rowsInserted);   
                            await cmd.ExecuteNonQueryAsync(); 
                        }

                        trans.Commit();
                        _logger?.Invoke($"[DB-OK] Đã nạp thành công {rowsInserted} dòng mới từ tệp {fileName}");
                        return rowsInserted; 
                    }
                    catch (Exception ex)
                    {
                        trans.Rollback();
                        _logger?.Invoke($"[DB-FAIL] Nạp dữ liệu thất bại cho tệp {fileName}: {ex}"); 
                        throw;
                    }
                }
            }
        }



        /// <summary>
        /// Xóa dữ liệu + lịch sử import của DANH SÁCH file cụ thể (chỉ file bị lỗi).
        /// Xử lý theo batch 50 file/lần để tránh query quá lớn.
        /// </summary>
        public async Task<(int deletedRows, int deletedHistory)> DeleteByFilePathsAsync(List<string> filePaths)
        {
            if (filePaths == null || filePaths.Count == 0) return (0, 0);

            int totalDeletedRows = 0;
            int totalDeletedHistory = 0;

            try
            {
                using (var conn = await CreateConnectionWithRetryAsync())
                {
                    // Xử lý theo batch 50 file/lần
                    const int batchSize = 50;
                    for (int batchStart = 0; batchStart < filePaths.Count; batchStart += batchSize)
                    {
                        var batch = filePaths.Skip(batchStart).Take(batchSize).ToList();

                        using (var trans = conn.BeginTransaction())
                        {
                            try
                            {
                                // Tạo danh sách tham số @p0, @p1, @p2...
                                var paramNames = new List<string>();
                                for (int i = 0; i < batch.Count; i++)
                                    paramNames.Add($"@p{i}");
                                string inClause = string.Join(", ", paramNames);

                                // 1. Xóa dữ liệu chính (toàn bộ dòng của file bị lỗi, cả NULL lẫn non-NULL)
                                string sqlDeleteData = $"DELETE FROM {TableData} WHERE FilePath IN ({inClause})";
                                using (var cmd = new SqlCommand(sqlDeleteData, conn, trans))
                                {
                                    cmd.CommandTimeout = 300;
                                    for (int i = 0; i < batch.Count; i++)
                                        cmd.Parameters.AddWithValue($"@p{i}", batch[i]);
                                    totalDeletedRows += await cmd.ExecuteNonQueryAsync();
                                }

                                // 2. Xóa lịch sử import
                                string sqlDeleteHistory = $"DELETE FROM {TableHistory} WHERE FilePath IN ({inClause})";
                                using (var cmd = new SqlCommand(sqlDeleteHistory, conn, trans))
                                {
                                    cmd.CommandTimeout = 120;
                                    for (int i = 0; i < batch.Count; i++)
                                        cmd.Parameters.AddWithValue($"@p{i}", batch[i]);
                                    totalDeletedHistory += await cmd.ExecuteNonQueryAsync();
                                }

                                trans.Commit();
                            }
                            catch
                            {
                                trans.Rollback();
                                throw;
                            }
                        }
                    }

                    _logger?.Invoke($"[REBUILD] Đã xóa {totalDeletedRows:N0} dòng dữ liệu và {totalDeletedHistory:N0} bản ghi lịch sử từ {filePaths.Count} file lỗi.");
                    return (totalDeletedRows, totalDeletedHistory);
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI-SQL] Xóa dữ liệu file lỗi thất bại: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Lấy thống kê tổng số file đã import thành công và tổng số dòng dữ liệu đã chèn.
        /// </summary>
        public async Task<(int totalFiles, int totalRows)> GetStatsAsync()
        {
            try
            {
                using (var conn = await CreateConnectionWithRetryAsync())
                {
                    string sql = $"SELECT COUNT(*), ISNULL(SUM(RowsInserted), 0) FROM {TableHistory} WHERE Status = 'Success'";
                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        cmd.CommandTimeout = 30;
                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                int files = reader.GetInt32(0);
                                int rows = reader.GetInt32(1);
                                return (files, rows);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI-SQL] Không lấy được thống kê dữ liệu: {ex.Message}");
            }
            return (0, 0);
        }
    }
}
