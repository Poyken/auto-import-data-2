using System;
using System.IO;
using System.Data;
using Xunit;
using ImportData.Services;

namespace ImportData.Tests
{
    public class ExcelServiceTests
    {
        private readonly ExcelService _excelService;
        private readonly string _workspaceDir;
        private readonly Xunit.Abstractions.ITestOutputHelper _output;

        public ExcelServiceTests(Xunit.Abstractions.ITestOutputHelper output)
        {
            _output = output;
            _excelService = new ExcelService(msg => Console.WriteLine(msg));

            string currentDir = AppDomain.CurrentDomain.BaseDirectory;
            while (currentDir != null && !Directory.Exists(Path.Combine(currentDir, "ExcelData")))
            {
                currentDir = Path.GetDirectoryName(currentDir);
            }
            _workspaceDir = currentDir ?? @"c:\Users\User Vinatech.DESKTOP-RJJSEQU\Desktop\auto-import-data-2";
        }

        [Fact]
        public void PrintColumns()
        {
            string[] files = {
                Path.Combine(_workspaceDir, "ExcelData", "2026-08-15", "1#15_14.52.13", "1#15_14.52.13.xlsx"),
                Path.Combine(_workspaceDir, "ExcelData", "2026-08-15", "2#15_15.05.01", "2#15_15.05.01.xlsx")
            };

            foreach (var file in files)
            {
                if (!File.Exists(file)) continue;

                _output.WriteLine($"\n=======================================================");
                _output.WriteLine($"INSPECTING FILE: {Path.GetFileName(file)}");
                _output.WriteLine($"=======================================================");

                var dt = _excelService.ReadExcelFile(file);
                if (dt == null)
                {
                    _output.WriteLine("FAILED TO READ EXCEL FILE!");
                    continue;
                }

                _output.WriteLine($"Total Rows: {dt.Rows.Count}, Total Columns: {dt.Columns.Count}\n");

                _output.WriteLine("--- COLUMN MAPPINGS TO SQL V2 ---");
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    string rawCol = dt.Columns[c].ColumnName;
                    string key = DatabaseService.GetSearchKey(rawCol);
                    bool mapped = DatabaseService.AliasToSqlColumnMap.TryGetValue(key, out string sqlCol);
                    string target = mapped ? sqlCol : "UNMAPPED";
                    _output.WriteLine($"Col[{c:D2}] Raw: '{rawCol}' => Key: '{key}' => SQL Target: '{target}'");
                }

                _output.WriteLine("\n--- SAMPLE DATA (FIRST 3 ROWS) ---");
                for (int r = 0; r < Math.Min(3, dt.Rows.Count); r++)
                {
                    _output.WriteLine($"\n[ROW {r + 1}]:");
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        string rawCol = dt.Columns[c].ColumnName;
                        string key = DatabaseService.GetSearchKey(rawCol);
                        DatabaseService.AliasToSqlColumnMap.TryGetValue(key, out string sqlCol);
                        object val = dt.Rows[r][c];
                        _output.WriteLine($"  {sqlCol ?? rawCol} ({rawCol}): '{val}'");
                    }
                }
            }
        }

        [Fact]
        public void Test_NormalizeColumnName_HiddenCharsAndNewlines()
        {
            string col1 = "Equipment Number\u200B";
            Assert.Equal("equipmentnumber", ExcelService.NormalizeColumnName(col1));

            string col2 = "Charge\nEndCurrent(mA)";
            Assert.Equal("chargeendcurrentma", ExcelService.NormalizeColumnName(col2));

            string col3 = "DischargeVoltage1_Time (mm：ss)";
            Assert.Equal("dischargevoltage1timemmss", ExcelService.NormalizeColumnName(col3));

            string col4 = "\uFEFFBarcode";
            Assert.Equal("barcode", ExcelService.NormalizeColumnName(col4));
        }

        [Fact]
        public void Test_ReadExcelFile_Format1_EquipmentNumber()
        {
            string filePath = Path.Combine(_workspaceDir, "ExcelData", "2026-08-15", "1#15_14.52.13", "1#15_14.52.13.xlsx");
            if (!File.Exists(filePath)) return;

            DataTable dt = _excelService.ReadExcelFile(filePath);
            if (dt != null)
            {
                Assert.True(_excelService.ValidateHeaders(dt));
                Assert.True(dt.Rows.Count > 0);
            }
        }

        [Fact]
        public void Test_ReadExcelFile_Format2_DevName()
        {
            string filePath = Path.Combine(_workspaceDir, "ExcelData", "2026-08-15", "2#15_15.05.01", "2#15_15.05.01.xlsx");
            if (!File.Exists(filePath)) return;

            DataTable dt = _excelService.ReadExcelFile(filePath);
            if (dt != null)
            {
                Assert.True(_excelService.ValidateHeaders(dt));
                Assert.True(dt.Rows.Count > 0);
            }
        }

        [Fact]
        public void Test_NormalizeColumnName_ExtremeInputs()
        {
            Assert.Equal(string.Empty, ExcelService.NormalizeColumnName(null!));
            Assert.Equal(string.Empty, ExcelService.NormalizeColumnName(""));

            string input1 = "\tEquipment   \n   Number_#!%";
            Assert.Equal("equipmentnumber#!", ExcelService.NormalizeColumnName(input1));

            string input2 = "Discharge\r/Voltage\\1_Time";
            Assert.Equal("dischargevoltage1time", ExcelService.NormalizeColumnName(input2));
        }

        [Fact]
        public void Test_ValidateHeaders_Thresholds()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Barcode");
            dt.Columns.Add("StartTime");
            Assert.False(_excelService.ValidateHeaders(dt));

            dt.Columns.Add("Unrelated1");
            dt.Columns.Add("Unrelated2");
            dt.Columns.Add("Unrelated3");
            Assert.False(_excelService.ValidateHeaders(dt));

            dt.Columns.Add("SorterNum");
            dt.Columns.Add("Slot");
            dt.Columns.Add("Position");
            dt.Columns.Add("Channel");
            dt.Columns.Add("Capacity");
            dt.Columns.Add("Capacitance");
            
            Assert.True(_excelService.ValidateHeaders(dt));
        }

        [Fact]
        public void Test_CleanEmptyCells_And_ReplaceDashesWithDbNull()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Col1", typeof(object));
            dt.Columns.Add("Col2", typeof(object));

            dt.Rows.Add("Value1", "Value2");
            dt.Rows.Add("---", "---");
            dt.Rows.Add("   ", "");
            dt.Rows.Add("Value3", "---");

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

            Assert.Equal(2, dt.Rows.Count);
            Assert.Equal("Value1", dt.Rows[0]["Col1"]);
            Assert.Equal("Value2", dt.Rows[0]["Col2"]);
            Assert.Equal("Value3", dt.Rows[1]["Col1"]);
            Assert.Equal(DBNull.Value, dt.Rows[1]["Col2"]);
        }

        [Fact]
        public void Test_ReadExcelFile_Chinese2Step_File()
        {
            string filePath = @"D:\ExcelData\2026-08-21\2#21_13.14.11\2#21_13.14.11.xlsx";
            if (!File.Exists(filePath)) return;

            DataTable dt = _excelService.ReadExcelFile(filePath);
            Assert.NotNull(dt);
            Assert.True(dt.Rows.Count > 0);
            Assert.True(_excelService.ValidateHeaders(dt));

            // Verify that CCDchg and Rest columns are mapped properly
            bool hasChannel = false;
            bool hasPosition = false;
            bool hasCcdchgCapacity = false;
            bool hasRestVoltage = false;

            for (int c = 0; c < dt.Columns.Count; c++)
            {
                string col = dt.Columns[c].ColumnName;
                string key = DatabaseService.GetSearchKey(col);
                if (DatabaseService.AliasToSqlColumnMap.TryGetValue(key, out string? sqlCol))
                {
                    if (sqlCol == "Channel") hasChannel = true;
                    if (sqlCol == "Position") hasPosition = true;
                    if (sqlCol == "CCDchg_Capacity_mAh") hasCcdchgCapacity = true;
                    if (sqlCol == "Rest_BeginVoltage_mV") hasRestVoltage = true;
                }
            }

            Assert.True(hasChannel, "Missing Channel column");
            Assert.True(hasPosition, "Missing Position column");
            Assert.True(hasCcdchgCapacity, "Missing CCDchg_Capacity_mAh column");
            Assert.True(hasRestVoltage, "Missing Rest_BeginVoltage_mV column");
        }

        [Fact]
        public void Test_ReadExcelFile_Chinese3Step_File()
        {
            string filePath = @"D:\ExcelData\2026-08-18\1#18_08.52.44\1#18_08.52.44.xlsx";
            if (!File.Exists(filePath)) return;

            DataTable dt = _excelService.ReadExcelFile(filePath);
            Assert.NotNull(dt);
            Assert.True(dt.Rows.Count > 0);
            Assert.True(_excelService.ValidateHeaders(dt));

            bool hasCccvChg = false;
            bool hasCcdchg = false;
            bool hasRest = false;

            for (int c = 0; c < dt.Columns.Count; c++)
            {
                string col = dt.Columns[c].ColumnName;
                string key = DatabaseService.GetSearchKey(col);
                if (DatabaseService.AliasToSqlColumnMap.TryGetValue(key, out string? sqlCol))
                {
                    if (sqlCol?.StartsWith("CCCVChg_") == true) hasCccvChg = true;
                    if (sqlCol?.StartsWith("CCDchg_") == true) hasCcdchg = true;
                    if (sqlCol?.StartsWith("Rest_") == true) hasRest = true;
                }
            }

            Assert.True(hasCccvChg, "Missing CCCVChg columns");
            Assert.True(hasCcdchg, "Missing CCDchg columns");
            Assert.True(hasRest, "Missing Rest columns");
        }

        [Fact]
        public void Test_ValidateHeaders_ChineseHeaders()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("通道");
            dt.Columns.Add("位置");
            dt.Columns.Add("电池ID");
            dt.Columns.Add("CCDchg_容量(mAh)");
            dt.Columns.Add("Rest_开始电压(mV)");

            Assert.True(_excelService.ValidateHeaders(dt));
        }

        [Fact]
        public void Test_ReadAll503FilesInDExcelData()
        {
            if (!Directory.Exists(@"D:\ExcelData")) return;

            var files = Directory.GetFiles(@"D:\ExcelData", "*.xlsx", SearchOption.AllDirectories)
                .Where(f => {
                    string name = Path.GetFileName(f);
                    return !name.StartsWith("~") && !name.Contains("~$") && !name.StartsWith("$");
                })
                .ToArray();

            int failedCount = 0;
            var failedFiles = new List<string>();

            foreach (var file in files)
            {
                var dt = _excelService.ReadExcelFile(file);
                if (dt == null || dt.Rows.Count == 0 || !_excelService.ValidateHeaders(dt))
                {
                    failedCount++;
                    failedFiles.Add(file);
                }
            }

            Assert.True(failedCount == 0, $"Failed reading {failedCount} files: {string.Join(", ", failedFiles.Take(5))}");
        }
    }
}
