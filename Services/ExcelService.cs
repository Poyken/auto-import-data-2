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
        /// Mảng RequiredHeaders: Danh sách tiêu đề cột nhận diện dữ liệu đo đạc (Cả Ver 1, Ver 2, tiếng Anh và tiếng Trung).
        /// </summary>
        internal static readonly string[] RequiredHeaders = {
            "EquipmentNumber", "DevName", "SorterNum", "SortNum", "TrayID", "StartTime", "BeginTime", "EndTime", 
            "WorkflowCode", "LotNo", "Barcode", "Slot", "Position", "Channel", "Capacity", "Capacitance", 
            "WorkType", "WorkstepTime", "StopReason", "BeginVoltage", "EndVoltage", "BeginCurrent", "EndCurrent",
            "BeginVoltageSD", "ChargeEndCurrent", "DischargeVoltage1", "DischargeVoltage1Time", 
            "DischargeVoltage2", "DischargeVoltage2Time", "DischargeBeginVoltage", "DischargeBeginCurrent", "NGInfo",
            // Tiếng Trung:
            "通道", "位置", "托盘", "托盘id", "电池", "电池id", "条码", "工作时间", "截止原因", "开始电压", "结束电压", 
            "开始时间", "结束时间", "开始端口电压", "开始电流", "结束电流", "结束端口电压", "容量", "电容", "电容值", "时长"
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
                            if (table.TableName.Contains("AggregateData", StringComparison.OrdinalIgnoreCase) || 
                                table.TableName.Contains("综合", StringComparison.OrdinalIgnoreCase))
                            {
                                selectedTable = table;
                                break;
                            }
                        }
                        if (selectedTable == null)
                        {
                            foreach (DataTable table in ds.Tables)
                            {
                                if (table.Columns.Count >= 10)
                                {
                                    selectedTable = table;
                                    break;
                                }
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
                                if (val.Contains("channel") || val.Contains("position") || val.Contains("barcode") || val.Contains("workstep") ||
                                    val.Contains("通道") || val.Contains("位置") || val.Contains("电池") || val.Contains("工作时间") || val.Contains("托盘") || val.Contains("条码"))
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
                            // Nhận diện Step Group từ hàng trên Header (nếu có hàng group)
                            string[] groupPrefixes = new string[selectedTable.Columns.Count];
                            if (headerRowIdx > 0)
                            {
                                int groupRowIdx = headerRowIdx - 1;
                                string currentGroup = "";
                                for (int c = 0; c < selectedTable.Columns.Count; c++)
                                {
                                    string groupVal = selectedTable.Rows[groupRowIdx][c]?.ToString()?.Trim() ?? "";
                                    if (!string.IsNullOrEmpty(groupVal))
                                    {
                                        string normGroup = groupVal.ToLower();
                                        if (normGroup.Contains("cccv") || normGroup.Contains("恒流恒压") || normGroup.Contains("充电"))
                                        {
                                            currentGroup = "CCCVChg";
                                        }
                                        else if (normGroup.Contains("ccd") || normGroup.Contains("恒流放电") || normGroup.Contains("放电"))
                                        {
                                            currentGroup = "CCDchg";
                                        }
                                        else if (normGroup.Contains("rest") || normGroup.Contains("搁置") || normGroup.Contains("静置"))
                                        {
                                            currentGroup = "Rest";
                                        }
                                        else if (normGroup.Contains("channel") || normGroup.Contains("通道") || normGroup.Contains("info") || normGroup.Contains("信息"))
                                        {
                                            currentGroup = "";
                                        }
                                    }
                                    groupPrefixes[c] = currentGroup;
                                }
                            }

                            for (int c = 0; c < selectedTable.Columns.Count; c++)
                            {
                                string rawColName = selectedTable.Rows[headerRowIdx][c]?.ToString()?.Trim() ?? "";
                                if (string.IsNullOrEmpty(rawColName)) rawColName = $"Col_{c}";

                                string prefix = groupPrefixes[c];
                                string normRaw = NormalizeColumnName(rawColName);
                                // Không gắn prefix cho các cột thông tin chung (Channel, Position, TrayID, Barcode)
                                bool isCommonCol = normRaw == "channel" || normRaw == "position" || normRaw == "trayid" || normRaw == "barcode" ||
                                                   normRaw == "通道" || normRaw == "位置" || normRaw == "托盘id" || normRaw == "托盘" || normRaw == "电池id" || normRaw == "条码";

                                string colName = (!string.IsNullOrEmpty(prefix) && !isCommonCol) ? $"{prefix}_{rawColName}" : rawColName;

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
                    string reqName = NormalizeColumnName(required);
                    if (!string.IsNullOrEmpty(reqName) && (colName.Contains(reqName) || reqName.Contains(colName)))
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
