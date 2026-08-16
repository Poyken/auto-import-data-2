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
    /// Hỗ trợ cả định dạng V1 cũ và V2 mới từ các máy đo khác nhau.
    /// </summary>
    public class ExcelService
    {
        private readonly Action<string> _logger;

        /// <summary>
        /// Mảng RequiredHeaders: Danh sách tiêu đề cột nhận diện dữ liệu đo đạc (Cả Ver 1 và Ver 2).
        /// </summary>
        internal static readonly string[] RequiredHeaders = {
            "EquipmentNumber", "DevName", "SorterNum", "SortNum", "TrayID", "StartTime", "BeginTime", "EndTime", 
            "WorkflowCode", "LotNo", "Barcode", "Slot", "Position", "Channel", "Capacity", "Capacitance", 
            "WorkType", "WorkstepTime", "StopReason", "BeginVoltage", "EndVoltage", "BeginCurrent", "EndCurrent",
            "BeginVoltageSD", "ChargeEndCurrent", "DischargeVoltage1", "DischargeVoltage1Time", 
            "DischargeVoltage2", "DischargeVoltage2Time", "DischargeBeginVoltage", "DischargeBeginCurrent", "NGInfo"
        };

        public ExcelService(Action<string> logger)
        {
            _logger = logger; 
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        internal static string NormalizeColumnName(string name)
        {
            if (string.IsNullOrEmpty(name)) return string.Empty;
            var chars = name.Where(c => !char.IsControl(c) && (int)c != 8203 && (int)c != 65279).ToArray();
            string clean = new string(chars).Trim();
            return clean.Replace(" ", "").Replace("_", "").Replace("(", "").Replace(")", "")
                        .Replace("-", "").Replace("/", "").Replace("\\", "")
                        .Replace(":", "").Replace("：", "").Replace(",", "").Replace("%", "").ToLower();
        }

        public DataTable ReadExcelFile(string filePath)
        {
            try 
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) 
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream)) 
                    {
                        var ds = reader.AsDataSet(new ExcelDataSetConfiguration() 
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = false }
                        });

                        if (ds.Tables.Count == 0) return null;

                        DataTable selectedTable = null;
                        foreach (DataTable table in ds.Tables)
                        {
                            if (table.TableName.Contains("AggregateData", StringComparison.OrdinalIgnoreCase) || table.Columns.Count >= 10)
                            {
                                selectedTable = table;
                                break;
                            }
                        }
                        if (selectedTable == null) selectedTable = ds.Tables[0];

                        // Dò dòng chứa Header (Hàng 1 hoặc Hàng 2)
                        int headerRowIdx = -1;
                        for (int r = 0; r < Math.Min(5, selectedTable.Rows.Count); r++)
                        {
                            int matchCount = 0;
                            for (int c = 0; c < selectedTable.Columns.Count; c++)
                            {
                                string val = selectedTable.Rows[r][c]?.ToString()?.ToLower() ?? "";
                                if (val.Contains("channel") || val.Contains("position") || val.Contains("barcode") || val.Contains("workstep"))
                                {
                                    matchCount++;
                                }
                            }
                            if (matchCount >= 2)
                            {
                                headerRowIdx = r;
                                break;
                            }
                        }

                        DataTable dt = new DataTable();
                        if (headerRowIdx >= 0)
                        {
                            for (int c = 0; c < selectedTable.Columns.Count; c++)
                            {
                                string colName = selectedTable.Rows[headerRowIdx][c]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(colName)) colName = $"Col_{c}";
                                string finalColName = colName;
                                int dup = 1;
                                while (dt.Columns.Contains(finalColName)) finalColName = $"{colName}_{dup++}";
                                dt.Columns.Add(finalColName);
                            }

                            for (int r = headerRowIdx + 1; r < selectedTable.Rows.Count; r++)
                            {
                                DataRow newRow = dt.NewRow();
                                for (int c = 0; c < selectedTable.Columns.Count; c++)
                                {
                                    newRow[c] = selectedTable.Rows[r][c];
                                }
                                dt.Rows.Add(newRow);
                            }
                        }
                        else
                        {
                            dt = selectedTable;
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

                                for (int j = 0; j < dt.Columns.Count; j++)
                                {
                                    var valStr = row[j]?.ToString();
                                    if (valStr == "---" || string.IsNullOrWhiteSpace(valStr))
                                    {
                                        row[j] = DBNull.Value;
                                    }
                                }
                            }
                            return dt;
                        }
                    } 
                } 
                return null; 
            }
            catch (Exception ex) 
            {
                _logger?.Invoke($"[LỖI-EXCEL] Không xử lý được tệp Excel: {ex.Message}"); 
                return null; 
            }
        }

        internal bool ValidateHeaders(DataTable dt)
        {
            if (dt == null || dt.Columns.Count < 3) return false; 

            int matchCount = 0;
            foreach (DataColumn dc in dt.Columns)
            {
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
            return matchCount >= 3;
        }

        public bool HasEsrColumns(string filePath)
        {
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        if (reader.Read())
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
            catch { }
            return false;
        }
    }
}
