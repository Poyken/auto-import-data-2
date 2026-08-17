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
        private const string TableData = "SortingDataImportExcel_V2";
        private const string TableHistory = "ExcelImportHistory_V2";

        // Số lần retry tối đa khi gặp lỗi SQL transient
        private const int MaxRetryCount = 3;
        // Delay cơ bản giữa các lần retry (sẽ nhân lên theo cấp số nhân)
        private static readonly int[] RetryDelaysMs = { 2000, 5000, 10000 };

        public static readonly string[] SqlColumns = {
            "EquipmentNumber", "Position", "Channel", "TrayID", "Barcode",
            "CCCVChg_WorkstepTime", "CCCVChg_StopReason", "CCCVChg_BeginVoltage_mV", "CCCVChg_EndVoltage_mV", "CCCVChg_BeginTime", "CCCVChg_EndTime", "CCCVChg_BeginDKVoltage_mV", "CCCVChg_BeginCurrent_mA", "CCCVChg_EndCurrent_mA", "CCCVChg_EndDKVoltage_mV",
            "CCDchg_WorkstepTime", "CCDchg_StopReason", "CCDchg_BeginVoltage_mV", "CCDchg_EndVoltage_mV", "CCDchg_BeginTime", "CCDchg_EndTime", "CCDchg_BeginCurrent_mA", "CCDchg_EndCurrent_mA", "CCDchg_Capacity_mAh", "CCDchg_Capacitance_F", "CCDchg_Capacitance1_F", "CCDchg_CapacitanceVoltage2_mV", "CCDchg_Capacitance2_F", "CCDchg_Capacitance3_F", "CCDchg_Capacitance4_F",
            "Rest_WorkstepTime", "Rest_StopReason", "Rest_BeginVoltage_mV", "Rest_EndVoltage_mV", "Rest_BeginTime", "Rest_EndTime", "Rest_BeginDKVoltage_mV",
            "FilePath", "ImportDate"
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
        /// Bảng alias: Ánh xạ từ tên cột Excel (đã chuẩn hóa) sang tên cột SQL V2.
        /// </summary>
        internal static readonly Dictionary<string, string> AliasToSqlColumnMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // --- Cột thiết bị & vị trí ---
            { "equipmentnumber", "EquipmentNumber" },
            { "devname", "EquipmentNumber" },
            { "sorternum", "EquipmentNumber" },
            { "sortnum", "EquipmentNumber" },

            { "position", "Position" },
            { "slot", "Position" },

            { "channel", "Channel" },

            { "trayid", "TrayID" },
            { "trayno", "TrayID" },
            { "tray", "TrayID" },
            { "traynumber", "TrayID" },
            { "traycode", "TrayID" },

            { "barcode", "Barcode" },
            { "lotno", "Barcode" },
            { "lotid", "Barcode" },
            { "cellid", "Barcode" },
            { "barcodelotno", "Barcode" },
            { "barcodeno", "Barcode" },
            { "sn", "Barcode" },
            { "serialnumber", "Barcode" },
            { "serialno", "Barcode" },

            // --- 1. CCCVChg ---
            { "worksteptime", "CCCVChg_WorkstepTime" },
            { "worksteptime0", "CCCVChg_WorkstepTime" },
            { "stopreason", "CCCVChg_StopReason" },
            { "stopreason0", "CCCVChg_StopReason" },
            { "beginvoltagemv", "CCCVChg_BeginVoltage_mV" },
            { "beginvoltagemv0", "CCCVChg_BeginVoltage_mV" },
            { "beginvoltage", "CCCVChg_BeginVoltage_mV" },
            { "chargebeginvoltage", "CCCVChg_BeginVoltage_mV" },
            { "endvoltagemv", "CCCVChg_EndVoltage_mV" },
            { "endvoltagemv0", "CCCVChg_EndVoltage_mV" },
            { "endvoltage", "CCCVChg_EndVoltage_mV" },
            { "chargeendvoltage", "CCCVChg_EndVoltage_mV" },
            { "begintime", "CCCVChg_BeginTime" },
            { "begintime0", "CCCVChg_BeginTime" },
            { "starttime", "CCCVChg_BeginTime" },
            { "endtime", "CCCVChg_EndTime" },
            { "endtime0", "CCCVChg_EndTime" },
            { "begindkvoltagemv", "CCCVChg_BeginDKVoltage_mV" },
            { "begindkvoltagemv0", "CCCVChg_BeginDKVoltage_mV" },
            { "begincurrentma", "CCCVChg_BeginCurrent_mA" },
            { "begincurrentma0", "CCCVChg_BeginCurrent_mA" },
            { "begincurrent", "CCCVChg_BeginCurrent_mA" },
            { "chargebegincurrent", "CCCVChg_BeginCurrent_mA" },
            { "endcurrentma", "CCCVChg_EndCurrent_mA" },
            { "endcurrentma0", "CCCVChg_EndCurrent_mA" },
            { "endcurrent", "CCCVChg_EndCurrent_mA" },
            { "chargeendcurrent", "CCCVChg_EndCurrent_mA" },
            { "enddkvoltagemv", "CCCVChg_EndDKVoltage_mV" },
            { "enddkvoltagemv0", "CCCVChg_EndDKVoltage_mV" },

            // --- 2. CCDchg ---
            { "worksteptime1", "CCDchg_WorkstepTime" },
            { "worksteptime_1", "CCDchg_WorkstepTime" },
            { "stopreason1", "CCDchg_StopReason" },
            { "stopreason_1", "CCDchg_StopReason" },
            { "beginvoltagemv1", "CCDchg_BeginVoltage_mV" },
            { "beginvoltagemv_1", "CCDchg_BeginVoltage_mV" },
            { "dischargebeginvoltage", "CCDchg_BeginVoltage_mV" },
            { "dischargebeginvoltagemv", "CCDchg_BeginVoltage_mV" },
            { "endvoltagemv1", "CCDchg_EndVoltage_mV" },
            { "endvoltagemv_1", "CCDchg_EndVoltage_mV" },
            { "dischargeendvoltage", "CCDchg_EndVoltage_mV" },
            { "dischargeendvoltagemv", "CCDchg_EndVoltage_mV" },
            { "begintime1", "CCDchg_BeginTime" },
            { "begintime_1", "CCDchg_BeginTime" },
            { "endtime1", "CCDchg_EndTime" },
            { "endtime_1", "CCDchg_EndTime" },
            { "begincurrentma1", "CCDchg_BeginCurrent_mA" },
            { "begincurrentma_1", "CCDchg_BeginCurrent_mA" },
            { "dischargebegincurrent", "CCDchg_BeginCurrent_mA" },
            { "dischargebegincurrentma", "CCDchg_BeginCurrent_mA" },
            { "endcurrentma1", "CCDchg_EndCurrent_mA" },
            { "endcurrentma_1", "CCDchg_EndCurrent_mA" },
            { "dischargeendcurrent", "CCDchg_EndCurrent_mA" },
            { "dischargeendcurrentma", "CCDchg_EndCurrent_mA" },

            { "capacitymah", "CCDchg_Capacity_mAh" },
            { "capacity", "CCDchg_Capacity_mAh" },
            { "capacitancef", "CCDchg_Capacitance_F" },
            { "capacitance", "CCDchg_Capacitance_F" },
            { "capacitance1f", "CCDchg_Capacitance1_F" },
            { "capacitance1", "CCDchg_Capacitance1_F" },
            { "capacitancevoltage2mv", "CCDchg_CapacitanceVoltage2_mV" },
            { "capacitancevoltage2", "CCDchg_CapacitanceVoltage2_mV" },
            { "capacitance2f", "CCDchg_Capacitance2_F" },
            { "capacitance2", "CCDchg_Capacitance2_F" },
            { "capacitance3f", "CCDchg_Capacitance3_F" },
            { "capacitance3", "CCDchg_Capacitance3_F" },
            { "capacitance4f", "CCDchg_Capacitance4_F" },
            { "capacitance4", "CCDchg_Capacitance4_F" },

            // --- 3. Rest ---
            { "worksteptime2", "Rest_WorkstepTime" },
            { "worksteptime_2", "Rest_WorkstepTime" },
            { "stopreason2", "Rest_StopReason" },
            { "stopreason_2", "Rest_StopReason" },
            { "beginvoltagemv2", "Rest_BeginVoltage_mV" },
            { "beginvoltagemv_2", "Rest_BeginVoltage_mV" },
            { "endvoltagemv2", "Rest_EndVoltage_mV" },
            { "endvoltagemv_2", "Rest_EndVoltage_mV" },
            { "begintime2", "Rest_BeginTime" },
            { "begintime_2", "Rest_BeginTime" },
            { "endtime2", "Rest_EndTime" },
            { "endtime_2", "Rest_EndTime" },
            { "begindkvoltagemv1", "Rest_BeginDKVoltage_mV" },
            { "begindkvoltagemv_1", "Rest_BeginDKVoltage_mV" },
            { "begindkvoltagemv2", "Rest_BeginDKVoltage_mV" },
            { "begindkvoltagemv_2", "Rest_BeginDKVoltage_mV" },

            // Metadata
            { "filepath", "FilePath" },
            { "importdate", "ImportDate" }
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

        private async Task<SqlConnection> CreateConnectionWithRetryAsync()
        {
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
            var finalConn = new SqlConnection(GetOptimizedConnectionString());
            await finalConn.OpenAsync();
            return finalConn;
        }

        internal static bool IsTransientErrorNumber(int errorNumber)
        {
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

        public async Task<bool> IsFileImportedAsync(string filePath) 
        {
            try
            {
                if (!File.Exists(filePath)) return false;
                
                long fileSize = new FileInfo(filePath).Length;

                string relativePath = filePath;
                if (!string.IsNullOrEmpty(_config.BaseFolder) && filePath.StartsWith(_config.BaseFolder, StringComparison.OrdinalIgnoreCase))
                {
                    relativePath = filePath.Substring(_config.BaseFolder.Length).TrimStart('\\', '/');
                }
                relativePath = relativePath.Replace('/', '\\');

                using (var conn = await CreateConnectionWithRetryAsync()) 
                {
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
                throw;
            }
        }

        public async Task<int> ExecuteImportBatchAsync(DataTable dt, string fileName, string filePath) 
        {
            if (dt == null || dt.Rows.Count == 0) return 0;

            string eqNum = fileName.StartsWith("1#") ? "1#" : (fileName.StartsWith("2#") ? "2#" : "Unknown");
            DateTime importDate = DateTime.Now;

            // Xây dựng DataTable chuẩn hóa khớp 100% các cột SQL V2
            DataTable dbTable = new DataTable();
            foreach (var col in SqlColumns)
            {
                if (col == "ImportDate") dbTable.Columns.Add(col, typeof(DateTime));
                else dbTable.Columns.Add(col, typeof(object));
            }

            // Ánh xạ động từ chỉ số cột Excel sang tên cột SQL V2
            var excelColToSqlCol = new Dictionary<int, string>();
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                string colName = dt.Columns[c].ColumnName;
                string searchKey = GetSearchKey(colName);
                if (AliasToSqlColumnMap.TryGetValue(searchKey, out string sqlCol))
                {
                    if (!excelColToSqlCol.ContainsKey(c))
                    {
                        excelColToSqlCol[c] = sqlCol;
                    }
                }
            }

            // Bản đồ vị trí cột Excel fallback cho getVal nếu không khớp alias
            Dictionary<string, int> colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                string cName = dt.Columns[c].ColumnName.Trim();
                if (!colMap.ContainsKey(cName)) colMap[cName] = c;
            }

            Func<string, DataRow, object> getVal = (cName, dtRow) => {
                if (colMap.TryGetValue(cName, out int idx)) {
                    var v = dtRow[idx];
                    return (v == null || v == DBNull.Value || string.IsNullOrWhiteSpace(v.ToString()) || v.ToString() == "---") ? DBNull.Value : v;
                }
                return DBNull.Value;
            };

            for (int r = 0; r < dt.Rows.Count; r++)
            {
                DataRow dr = dbTable.NewRow();
                DataRow sourceRow = dt.Rows[r];

                // 1. Điền thông tin cố định / mặc định
                dr["EquipmentNumber"] = eqNum;
                dr["FilePath"] = filePath;
                dr["ImportDate"] = importDate;

                // 2. Điền giá trị từ Excel theo ánh xạ động (Smart Mapping)
                foreach (var kvp in excelColToSqlCol)
                {
                    int excelColIdx = kvp.Key;
                    string sqlColName = kvp.Value;
                    var val = sourceRow[excelColIdx];
                    if (val != null && val != DBNull.Value && !string.IsNullOrWhiteSpace(val.ToString()) && val.ToString() != "---")
                    {
                        if (dr[sqlColName] == DBNull.Value || (sqlColName == "EquipmentNumber" && val.ToString() != "Unknown"))
                        {
                            dr[sqlColName] = val;
                        }
                    }
                }

                // 3. Fallback điền bổ sung nếu cột vẫn là DBNull (đặc biệt các cột legacy)
                if (dr["Position"] == DBNull.Value) dr["Position"] = getVal("Position", sourceRow);
                if (dr["Channel"] == DBNull.Value) dr["Channel"] = getVal("Channel", sourceRow);
                if (dr["TrayID"] == DBNull.Value) dr["TrayID"] = getVal("TrayID", sourceRow);
                if (dr["Barcode"] == DBNull.Value) dr["Barcode"] = getVal("barcode", sourceRow);

                if (dr["CCCVChg_WorkstepTime"] == DBNull.Value) dr["CCCVChg_WorkstepTime"] = getVal("WorkstepTime", sourceRow);
                if (dr["CCCVChg_StopReason"] == DBNull.Value) dr["CCCVChg_StopReason"] = getVal("StopReason", sourceRow);
                if (dr["CCCVChg_BeginVoltage_mV"] == DBNull.Value) dr["CCCVChg_BeginVoltage_mV"] = getVal("BeginVoltage(mV)", sourceRow);
                if (dr["CCCVChg_EndVoltage_mV"] == DBNull.Value) dr["CCCVChg_EndVoltage_mV"] = getVal("EndVoltage(mV)", sourceRow);
                if (dr["CCCVChg_BeginTime"] == DBNull.Value) dr["CCCVChg_BeginTime"] = getVal("BeginTime", sourceRow);
                if (dr["CCCVChg_EndTime"] == DBNull.Value) dr["CCCVChg_EndTime"] = getVal("EndTime", sourceRow);

                if (dr["CCDchg_WorkstepTime"] == DBNull.Value) dr["CCDchg_WorkstepTime"] = getVal("WorkstepTime_1", sourceRow);
                if (dr["CCDchg_StopReason"] == DBNull.Value) dr["CCDchg_StopReason"] = getVal("StopReason_1", sourceRow);
                if (dr["CCDchg_BeginVoltage_mV"] == DBNull.Value) dr["CCDchg_BeginVoltage_mV"] = getVal("BeginVoltage(mV)_1", sourceRow);
                if (dr["CCDchg_EndVoltage_mV"] == DBNull.Value) dr["CCDchg_EndVoltage_mV"] = getVal("EndVoltage(mV)_1", sourceRow);
                if (dr["CCDchg_BeginTime"] == DBNull.Value) dr["CCDchg_BeginTime"] = getVal("BeginTime_1", sourceRow);
                if (dr["CCDchg_EndTime"] == DBNull.Value) dr["CCDchg_EndTime"] = getVal("EndTime_1", sourceRow);
                if (dr["CCDchg_Capacity_mAh"] == DBNull.Value) dr["CCDchg_Capacity_mAh"] = getVal("Capacity(mAh)", sourceRow);
                if (dr["CCDchg_Capacitance_F"] == DBNull.Value) dr["CCDchg_Capacitance_F"] = getVal("Capacitance(F)", sourceRow);

                dbTable.Rows.Add(dr);
            }

            using (var conn = await CreateConnectionWithRetryAsync()) 
            {
                using (var trans = conn.BeginTransaction())
                {
                    try
                    {
                        // 1. TẠO BẢNG TẠM TRONG SQL V2
                        string tempTableName = $"#TempImport_{Guid.NewGuid().ToString("N")}";
                        string createTempSql = $"SELECT TOP 0 * INTO {tempTableName} FROM {TableData}";
                        using (var cmdCreate = new SqlCommand(createTempSql, conn, trans)) { await cmdCreate.ExecuteNonQueryAsync(); }

                        // 2. BULK COPY VÀO BẢNG TẠM V2
                        using (var bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, trans))
                        {
                            bulkCopy.DestinationTableName = tempTableName;
                            bulkCopy.BatchSize = 1000;
                            bulkCopy.BulkCopyTimeout = 120;

                            foreach (DataColumn c in dbTable.Columns)
                            {
                                bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);
                            }
                            await bulkCopy.WriteToServerAsync(dbTable);
                        }

                        // 3. MERGE DỮ LIỆU TỪ BẢNG TẠM SANG BẢNG CHÍNH V2
                        string mergeSql = $@"
                            WITH UniqueSrc AS (
                                SELECT *, ROW_NUMBER() OVER (
                                    PARTITION BY EquipmentNumber, Position, Channel, FilePath
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
                                WHERE dest.EquipmentNumber = src.EquipmentNumber
                                AND dest.Position = src.Position
                                AND dest.Channel = src.Channel
                                AND dest.FilePath = src.FilePath
                            )";
                        
                        int rowsInserted = 0;
                        using (var cmdMerge = new SqlCommand(mergeSql, conn, trans))
                        {
                            cmdMerge.CommandTimeout = 300;
                            rowsInserted = await cmdMerge.ExecuteNonQueryAsync();
                        }

                        // 4. GHI LỊCH SỬ NẠP FILE V2
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
                        _logger?.Invoke($"[DB-OK] Đã nạp thành công {rowsInserted} dòng mới V2 từ tệp {fileName}");
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

        public async Task<(int deletedRows, int deletedHistory)> DeleteByFilePathsAsync(List<string> filePaths)
        {
            if (filePaths == null || filePaths.Count == 0) return (0, 0);

            int totalDeletedRows = 0;
            int totalDeletedHistory = 0;

            try
            {
                using (var conn = await CreateConnectionWithRetryAsync())
                {
                    const int batchSize = 50;
                    for (int batchStart = 0; batchStart < filePaths.Count; batchStart += batchSize)
                    {
                        var batch = filePaths.Skip(batchStart).Take(batchSize).ToList();

                        using (var trans = conn.BeginTransaction())
                        {
                            try
                            {
                                var paramNames = new List<string>();
                                for (int i = 0; i < batch.Count; i++)
                                    paramNames.Add($"@p{i}");
                                string inClause = string.Join(", ", paramNames);

                                string sqlDeleteData = $"DELETE FROM {TableData} WHERE FilePath IN ({inClause})";
                                using (var cmd = new SqlCommand(sqlDeleteData, conn, trans))
                                {
                                    cmd.CommandTimeout = 300;
                                    for (int i = 0; i < batch.Count; i++)
                                        cmd.Parameters.AddWithValue($"@p{i}", batch[i]);
                                    totalDeletedRows += await cmd.ExecuteNonQueryAsync();
                                }

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

                    _logger?.Invoke($"[REBUILD] Đã xóa {totalDeletedRows:N0} dòng dữ liệu và {totalDeletedHistory:N0} bản ghi lịch sử V2 từ {filePaths.Count} file lỗi.");
                    return (totalDeletedRows, totalDeletedHistory);
                }
            }
            catch (Exception ex)
            {
                _logger?.Invoke($"[LỖI-SQL] Xóa dữ liệu file lỗi thất bại: {ex.Message}");
                throw;
            }
        }

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
