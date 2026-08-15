using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace ImportData.Services 
{
    /// <summary>
    /// Lớp ExcelService: Chuyên đọc và phân giải nội dung tệp Excel từ máy đo.
    /// Chuyển đổi dữ liệu từ tệp vật lý sang bảng DataTable trong bộ nhớ RAM để SQL có thể nuốt được.
    /// Hỗ trợ 2 định dạng header khác nhau từ các máy đo khác nhau.
    /// </summary>
    public class ExcelService
    {
        private readonly Action<string> _logger; // Hàm log để bắn lỗi ra màn hình chính.

        /// <summary>
        /// Mảng RequiredHeaders: Danh sách tiêu đề cột bắt buộc phải có trong file Excel máy đo.
        /// Dùng để nhận diện file có phải là dữ liệu đo đạc không.
        /// Bao gồm alias từ cả 2 format (Equipment Number + DevName).
        /// </summary>
        internal static readonly string[] RequiredHeaders = {
            "EquipmentNumber", "DevName", "SorterNum", "SortNum", "StartTime", "WorkflowCode", "LotNo",
            "Barcode", "Slot", "Position", "Channel", "Capacity", "Capacitance", 
            "BeginVoltageSD", "ChargeEndCurrent", "EndVoltage", "EndCurrent", "DischargeVoltage1", 
            "DischargeVoltage1Time", "DischargeVoltage2", "DischargeVoltage2Time", "DischargeBeginVoltage", "DischargeBeginCurrent", 
            "NGInfo", "EndTime"
        };

        // Hàm khởi tạo ExcelService.
        public ExcelService(Action<string> logger)
        {
            _logger = logger; 
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        /// <summary>
        /// Chuẩn hóa tên cột Excel: Loại bỏ ký tự ẩn (newline, Zero-Width Space, BOM), 
        /// khoảng trắng, dấu đặc biệt để so sánh chính xác.
        /// </summary>
        internal static string NormalizeColumnName(string name)
        {
            if (string.IsNullOrEmpty(name)) return string.Empty;
            // Loại bỏ ký tự điều khiển (newline \n, carriage return \r, tab...), 
            // Zero-Width Space (\u200B = 8203), BOM (\uFEFF = 65279)
            var chars = name.Where(c => !char.IsControl(c) && (int)c != 8203 && (int)c != 65279).ToArray();
            string clean = new string(chars).Trim();
            return clean.Replace(" ", "").Replace("_", "").Replace("(", "").Replace(")", "")
                        .Replace("-", "").Replace("/", "").Replace("\\", "")
                        .Replace(":", "").Replace("：", "").Replace(",", "").Replace("%", "").ToLower();
        }

        /// <summary>
        /// Đọc nội dung tệp Excel vật lý và chuyển thành DataTable.
        /// Sử dụng FileShare.ReadWrite để không làm gián đoạn máy đo.
        /// </summary>
        public DataTable ReadExcelFile(string filePath)
        {
            try 
            {
                DataTable dt; 
                
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration() 
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        dt = result.Tables[0]; 
                    } 
                } 

                if (dt != null) 
                {
                    if (!ValidateHeaders(dt))
                    {
                        _logger?.Invoke($"[LỖI-EXCEL] Cấu trúc tệp không đúng: {Path.GetFileName(filePath)}"); 
                        return null; 
                    }

                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                    {
                        DataRow row = dt.Rows[i];
                        bool isEmpty = true;
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            var val = row[j];
                            if (val != null && !string.IsNullOrWhiteSpace(val.ToString()) && val.ToString() != "---")
                            {
                                isEmpty = false;
                                break;
                            }
                        }
                        
                        if (isEmpty) 
                        {
                            dt.Rows.RemoveAt(i);
                            continue;
                        }

                        // Chuyển đổi "---", rỗng hoặc chỉ có khoảng trắng thành DBNull cho các cột để tránh lỗi ép kiểu khi insert SQL
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            var valStr = row[j]?.ToString();
                            if (valStr == "---" || string.IsNullOrWhiteSpace(valStr))
                            {
                                row[j] = DBNull.Value;
                            }
                        }
                    }
                }
                return dt; 
            }
            catch (Exception ex) 
            {
                _logger?.Invoke($"[LỖI-EXCEL] Không xử lý được tệp Excel: {ex.Message}"); 
                return null; 
            }
        }

        /// <summary>
        /// Kiểm thử tiêu đề cột: Đảm bảo tệp Excel có cấu trúc khớp với dữ liệu máy đo.
        /// Hỗ trợ cả 2 format header (Equipment Number / DevName).
        /// Sử dụng NormalizeColumnName để loại bỏ ký tự ẩn, newline trước khi so sánh.
        /// </summary>
        internal bool ValidateHeaders(DataTable dt)
        {
            if (dt == null || dt.Columns.Count < 5) return false; 

            int matchCount = 0;
            foreach (DataColumn dc in dt.Columns)
            {
                // Chuẩn hóa tên cột: Loại bỏ ký tự ẩn, newline, dấu đặc biệt
                string colName = NormalizeColumnName(dc.ColumnName);
                
                foreach (var required in RequiredHeaders)
                {
                    string reqName = required.ToLower().Replace("_", "");
                    if (colName.Contains(reqName) || reqName.Contains(colName))
                    {
                        matchCount++;
                        break;
                    }
                }
            }

            // Nếu khớp được trên 8 cột tiêu đề quan trọng thì coi như đúng định dạng máy đo.
            return matchCount >= 8;
        }

        /// <summary>
        /// Kiểm tra nhanh xem file Excel có cột ESR, OCV hoặc ESRTime hay không.
        /// Chỉ đọc hàng tiêu đề đầu tiên để tối ưu tốc độ.
        /// </summary>
        public bool HasEsrColumns(string filePath)
        {
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        if (reader.Read()) // Đọc hàng đầu tiên
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                string colName = NormalizeColumnName(reader.GetValue(i)?.ToString() ?? "");
                                if (colName.Contains("esrm") || colName.Contains("esr") || colName.Contains("ocv"))
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            catch { /* Bỏ qua nếu có lỗi truy cập file */ }
            return false;
        }
    }
}
