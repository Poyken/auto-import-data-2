using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using Xunit;
using ImportData.Services;
using ImportData.Core;
using Microsoft.Data.SqlClient;

namespace ImportData.Tests
{
    public class DatabaseServiceTests
    {
        private readonly Xunit.Abstractions.ITestOutputHelper _output;

        public DatabaseServiceTests(Xunit.Abstractions.ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void Test_GetSearchKey_Normalization()
        {
            Assert.Equal("equipmentnumber", DatabaseService.GetSearchKey("Equipment Number\u200B"));
            Assert.Equal("chargeendcurrentma", DatabaseService.GetSearchKey("Charge\nEndCurrent(mA)"));
            Assert.Equal("dischargevoltage1timemmss", DatabaseService.GetSearchKey("DischargeVoltage1_Time (mm：ss)"));
            Assert.Equal("barcode", DatabaseService.GetSearchKey("\uFEFFBarcode"));
        }

        [Theory]
        [InlineData("equipmentnumber", "EquipmentNumber")]
        [InlineData("position", "Position")]
        [InlineData("channel", "Channel")]
        [InlineData("trayid", "TrayID")]
        [InlineData("trayno", "TrayID")]
        [InlineData("barcode", "Barcode")]
        [InlineData("lotno", "Barcode")]
        [InlineData("worksteptime", "CCCVChg_WorkstepTime")]
        [InlineData("worksteptime1", "CCDchg_WorkstepTime")]
        [InlineData("worksteptime2", "Rest_WorkstepTime")]
        [InlineData("dischargebeginvoltage", "CCDchg_BeginVoltage_mV")]
        [InlineData("endvoltagemv2", "Rest_EndVoltage_mV")]
        [InlineData("capacitymah", "CCDchg_Capacity_mAh")]
        [InlineData("capacitancef", "CCDchg_Capacitance_F")]
        [InlineData("通道", "Channel")]
        [InlineData("位置", "Position")]
        [InlineData("托盘id", "TrayID")]
        [InlineData("电池id", "Barcode")]
        [InlineData("cccvchg工作时间", "CCCVChg_WorkstepTime")]
        [InlineData("ccdchg工作时间", "CCDchg_WorkstepTime")]
        [InlineData("rest工作时间", "Rest_WorkstepTime")]
        [InlineData("ccdchg容量mah", "CCDchg_Capacity_mAh")]
        [InlineData("ccdchg电容f", "CCDchg_Capacitance_F")]
        [InlineData("rest开始端口电压mv", "Rest_BeginDKVoltage_mV")]
        public void Test_AliasToSqlColumnMap_CorrectMapping(string alias, string expectedSqlColumn)
        {
            bool hasMapping = DatabaseService.AliasToSqlColumnMap.TryGetValue(alias, out string? sqlCol);
            Assert.True(hasMapping, $"Thiếu ánh xạ cho alias: '{alias}'");
            Assert.Equal(expectedSqlColumn, sqlCol);
        }

        [Theory]
        [InlineData("20260531191549-6#A2.xlsx", "2026-05-31 19:15:49")]
        [InlineData("20260601000502-6#A2.xlsx", "2026-06-01 00:05:02")]
        [InlineData("Copy_20260601000502-6#A2.xlsx", "2026-06-01 00:05:02")]
        [InlineData("6#A2_20260601000502.xlsx", "2026-06-01 00:05:02")]
        public void Test_FileName_TimestampParsing(string filename, string expectedDateTimeStr)
        {
            DateTime expectedTime = DateTime.Parse(expectedDateTimeStr);
            
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filename);
            var match = System.Text.RegularExpressions.Regex.Match(fileNameWithoutExt, @"\d{14}");
            Assert.True(match.Success);
            
            bool success = DateTime.TryParseExact(match.Value, "yyyyMMddHHmmss", 
                System.Globalization.CultureInfo.InvariantCulture, 
                System.Globalization.DateTimeStyles.None, 
                out DateTime parsedTime);
                
            Assert.True(success);
            Assert.Equal(expectedTime, parsedTime);
        }

        [Theory]
        [InlineData("no_numbers_in_this_file.xlsx")]
        [InlineData("only_12_nums_123456789012.xlsx")]
        [InlineData("invalid_date_99999999999999.xlsx")]
        public void Test_FileName_TimestampParsing_EdgeCases(string filename)
        {
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filename);
            var match = System.Text.RegularExpressions.Regex.Match(fileNameWithoutExt, @"\d{14}");
            
            if (match.Success)
            {
                bool success = DateTime.TryParseExact(match.Value, "yyyyMMddHHmmss", 
                    System.Globalization.CultureInfo.InvariantCulture, 
                    System.Globalization.DateTimeStyles.None, 
                    out DateTime _);
                Assert.False(success, $"Should fail parsing invalid date string: {match.Value}");
            }
            else
            {
                Assert.False(match.Success, "Should not find any 14-digit match");
            }
        }

        [Theory]
        [InlineData(-2, true)]
        [InlineData(20, true)]
        [InlineData(64, true)]
        [InlineData(233, true)]
        [InlineData(10054, true)]
        [InlineData(-1, true)]
        [InlineData(258, true)]
        [InlineData(102, false)]
        [InlineData(547, false)]
        [InlineData(2627, false)]
        public void Test_IsTransientErrorNumber(int errorCode, bool expectedIsTransient)
        {
            bool actual = DatabaseService.IsTransientErrorNumber(errorCode);
            Assert.Equal(expectedIsTransient, actual);
        }

        [Fact]
        public void CrossReference_D_ExcelData_With_SQL()
        {
            if (!Directory.Exists(@"D:\ExcelData")) return;

            var excelService = new ExcelService(msg => _output.WriteLine(msg));
            var allFiles = Directory.GetFiles(@"D:\ExcelData", "*.xlsx", SearchOption.AllDirectories);
            _output.WriteLine($"Total Excel files found in D:\\ExcelData: {allFiles.Length}");

            var allUniqueHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var headerToFileCount = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var sampleValuesByHeader = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            var filesWithBarcode = new List<string>();
            var filesWithTrayId = new List<string>();
            var filesWithWorkType = new List<string>();
            var formatTypes = new Dictionary<string, int>();

            int filesAnalyzed = 0;
            foreach (var file in allFiles)
            {
                var dt = excelService.ReadExcelFile(file);
                if (dt == null)
                {
                    _output.WriteLine($"[CANNOT READ] {file}");
                    continue;
                }
                filesAnalyzed++;

                bool hasBarcodeVal = false;
                bool hasTrayIdVal = false;

                var headerSig = string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).OrderBy(c => c));
                if (!formatTypes.ContainsKey(headerSig)) formatTypes[headerSig] = 0;
                formatTypes[headerSig]++;

                foreach (DataColumn col in dt.Columns)
                {
                    string colName = col.ColumnName;
                    allUniqueHeaders.Add(colName);
                    if (!headerToFileCount.ContainsKey(colName))
                        headerToFileCount[colName] = 0;
                    headerToFileCount[colName]++;

                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        var v = dt.Rows[r][colName];
                        if (v != null && v != DBNull.Value && !string.IsNullOrWhiteSpace(v.ToString()) && v.ToString() != "---")
                        {
                            if (!sampleValuesByHeader.ContainsKey(colName))
                            {
                                sampleValuesByHeader[colName] = v;
                            }
                            if (colName.Equals("barcode", StringComparison.OrdinalIgnoreCase)) hasBarcodeVal = true;
                            if (colName.Equals("trayid", StringComparison.OrdinalIgnoreCase)) hasTrayIdVal = true;
                        }
                    }
                }

                if (hasBarcodeVal) filesWithBarcode.Add(Path.GetFileName(file));
                if (hasTrayIdVal) filesWithTrayId.Add(Path.GetFileName(file));
                if (dt.Columns.Contains("WorkType")) filesWithWorkType.Add(Path.GetFileName(file));
            }

            _output.WriteLine($"\n=======================================================");
            _output.WriteLine($"ANALYSIS OF HEADERS IN D:\\ExcelData (All {filesAnalyzed} files)");
            _output.WriteLine($"=======================================================");
            
            foreach (var header in allUniqueHeaders.OrderBy(h => h))
            {
                string key = DatabaseService.GetSearchKey(header);
                bool mapped = DatabaseService.AliasToSqlColumnMap.TryGetValue(key, out string sqlTarget);
                sampleValuesByHeader.TryGetValue(header, out object sampleVal);

                string status = mapped ? $"-> SQL: [{sqlTarget}]" : "-> [UNMAPPED / MISSING!]";
                _output.WriteLine($"Header: '{header}' (Key: '{key}') {status} | Sample Val: '{sampleVal}' | Appears in {headerToFileCount[header]}/{filesAnalyzed} files");
            }

            _output.WriteLine($"\nDistinct Header Structures found: {formatTypes.Count}");
            int sigIndex = 1;
            foreach (var kvp in formatTypes)
            {
                _output.WriteLine($"\n[Format Variant {sigIndex++}] ({kvp.Value} files):");
                _output.WriteLine($"Columns: {kvp.Key}");
            }

            _output.WriteLine($"\nFiles with NON-EMPTY Barcode: {filesWithBarcode.Count}/{filesAnalyzed}");
            _output.WriteLine($"Files with NON-EMPTY TrayID: {filesWithTrayId.Count}/{filesAnalyzed}");
            _output.WriteLine($"Files with WorkType columns: {filesWithWorkType.Count}/{filesAnalyzed}");
        }

        [Fact]
        public void Check_D_ExcelData_ImportCompleteness_In_SQL()
        {
            var config = new AppConfig();
            config.Load();

            using (var conn = new SqlConnection(config.ConnectionString))
            {
                conn.Open();
                _output.WriteLine("=== AUDIT D:\\ExcelData IN SQL SERVER ===");

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT Status, COUNT(*) as Cnt, SUM(RowsInserted) as TotalRows 
                        FROM ExcelImportHistory_V2 
                        WHERE FilePath LIKE 'D:\ExcelData%' 
                        GROUP BY Status";
                    using (var r = cmd.ExecuteReader())
                    {
                        _output.WriteLine("\n--- History Status for D:\\ExcelData ---");
                        while (r.Read())
                        {
                            _output.WriteLine($"Status: {r["Status"]} | Files: {r["Cnt"]} | TotalRows: {r["TotalRows"]}");
                        }
                    }

                    cmd.CommandText = @"
                        SELECT 
                            COUNT(*) as TotalRows,
                            COUNT(DISTINCT FilePath) as TotalFiles,
                            COUNT(EquipmentNumber) as EquipmentNumber,
                            COUNT(Position) as Position,
                            COUNT(Channel) as Channel,
                            COUNT(TrayID) as TrayID,
                            COUNT(Barcode) as Barcode,
                            COUNT(CCCVChg_WorkstepTime) as CCCVChg_WorkstepTime,
                            COUNT(CCCVChg_StopReason) as CCCVChg_StopReason,
                            COUNT(CCCVChg_BeginVoltage_mV) as CCCVChg_BeginVoltage_mV,
                            COUNT(CCCVChg_EndVoltage_mV) as CCCVChg_EndVoltage_mV,
                            COUNT(CCCVChg_BeginTime) as CCCVChg_BeginTime,
                            COUNT(CCCVChg_EndTime) as CCCVChg_EndTime,
                            COUNT(CCCVChg_BeginDKVoltage_mV) as CCCVChg_BeginDKVoltage_mV,
                            COUNT(CCCVChg_BeginCurrent_mA) as CCCVChg_BeginCurrent_mA,
                            COUNT(CCCVChg_EndCurrent_mA) as CCCVChg_EndCurrent_mA,
                            COUNT(CCCVChg_EndDKVoltage_mV) as CCCVChg_EndDKVoltage_mV,
                            COUNT(CCDchg_WorkstepTime) as CCDchg_WorkstepTime,
                            COUNT(CCDchg_StopReason) as CCDchg_StopReason,
                            COUNT(CCDchg_BeginVoltage_mV) as CCDchg_BeginVoltage_mV,
                            COUNT(CCDchg_EndVoltage_mV) as CCDchg_EndVoltage_mV,
                            COUNT(CCDchg_BeginTime) as CCDchg_BeginTime,
                            COUNT(CCDchg_EndTime) as CCDchg_EndTime,
                            COUNT(CCDchg_BeginCurrent_mA) as CCDchg_BeginCurrent_mA,
                            COUNT(CCDchg_EndCurrent_mA) as CCDchg_EndCurrent_mA,
                            COUNT(CCDchg_Capacity_mAh) as CCDchg_Capacity_mAh,
                            COUNT(CCDchg_Capacitance_F) as CCDchg_Capacitance_F,
                            COUNT(CCDchg_Capacitance1_F) as CCDchg_Capacitance1_F,
                            COUNT(CCDchg_CapacitanceVoltage2_mV) as CCDchg_CapacitanceVoltage2_mV,
                            COUNT(CCDchg_Capacitance2_F) as CCDchg_Capacitance2_F,
                            COUNT(CCDchg_Capacitance3_F) as CCDchg_Capacitance3_F,
                            COUNT(CCDchg_Capacitance4_F) as CCDchg_Capacitance4_F,
                            COUNT(Rest_WorkstepTime) as Rest_WorkstepTime,
                            COUNT(Rest_StopReason) as Rest_StopReason,
                            COUNT(Rest_BeginVoltage_mV) as Rest_BeginVoltage_mV,
                            COUNT(Rest_EndVoltage_mV) as Rest_EndVoltage_mV,
                            COUNT(Rest_BeginTime) as Rest_BeginTime,
                            COUNT(Rest_EndTime) as Rest_EndTime,
                            COUNT(Rest_BeginDKVoltage_mV) as Rest_BeginDKVoltage_mV
                        FROM SortingDataImportExcel_V2
                        WHERE FilePath LIKE 'D:\ExcelData%'";
                    using (var r = cmd.ExecuteReader())
                    {
                        if (r.Read())
                        {
                            _output.WriteLine("\n--- Column Non-Null Counts for D:\\ExcelData ---");
                            long total = Convert.ToInt64(r["TotalRows"]);
                            _output.WriteLine($"Total Rows in DB: {total}");
                            _output.WriteLine($"Total Files in DB: {r["TotalFiles"]}");

                            for (int i = 2; i < r.FieldCount; i++)
                            {
                                string col = r.GetName(i);
                                long cnt = Convert.ToInt64(r.GetValue(i));
                                double pct = total > 0 ? (cnt * 100.0 / total) : 0;
                                string warning = cnt == 0 ? " [ALL NULL 0%]" : (cnt < total ? $" [PARTIAL {pct:F1}%]" : " [FULL 100%]");
                                _output.WriteLine($"  {col,-32}: {cnt,7} / {total} ({pct,5:F1}%){warning}");
                            }
                        }
                    }
                }
            }
        }

        [Fact]
        public async Task Test_Import_Chinese2Step_2_21_13_14_11()
        {
            string filePath = @"D:\ExcelData\2026-08-21\2#21_13.14.11\2#21_13.14.11.xlsx";
            if (!File.Exists(filePath)) return;

            var config = new AppConfig();
            config.Load();
            var dbService = new DatabaseService(config, msg => _output.WriteLine(msg));
            var excelService = new ExcelService(msg => _output.WriteLine(msg));

            // 1. Read Excel
            var dt = excelService.ReadExcelFile(filePath);
            Assert.NotNull(dt);
            Assert.True(dt.Rows.Count > 0);
            _output.WriteLine($"Read {dt.Rows.Count} rows from {filePath}");

            // 2. Import into SQL Server
            int rowsInserted = await dbService.ExecuteImportBatchAsync(dt, Path.GetFileName(filePath), filePath);
            _output.WriteLine($"Rows inserted: {rowsInserted}");
            Assert.True(rowsInserted > 0);

            // 3. Verify in SQL Server
            using (var conn = new SqlConnection(config.ConnectionString))
            {
                await conn.OpenAsync();
                
                // Verify History
                using (var cmd = new SqlCommand("SELECT Status, RowsInserted FROM ExcelImportHistory_V2 WHERE FilePath = @path", conn))
                {
                    cmd.Parameters.AddWithValue("@path", filePath);
                    using (var r = await cmd.ExecuteReaderAsync())
                    {
                        Assert.True(await r.ReadAsync(), "History record not found!");
                        Assert.Equal("Success", r["Status"]?.ToString());
                        Assert.Equal(rowsInserted, Convert.ToInt32(r["RowsInserted"]));
                    }
                }

                // Verify Data rows
                using (var cmd = new SqlCommand("SELECT TOP 5 Position, Channel, CCDchg_Capacity_mAh, CCDchg_BeginVoltage_mV, CCDchg_EndVoltage_mV, Rest_BeginVoltage_mV FROM SortingDataImportExcel_V2 WHERE FilePath = @path", conn))
                {
                    cmd.Parameters.AddWithValue("@path", filePath);
                    using (var r = await cmd.ExecuteReaderAsync())
                    {
                        int rowCount = 0;
                        while (await r.ReadAsync())
                        {
                            rowCount++;
                            _output.WriteLine($"Sample Row: Pos={r["Position"]}, Ch={r["Channel"]}, Cap={r["CCDchg_Capacity_mAh"]}, CCDchg_BeginV={r["CCDchg_BeginVoltage_mV"]}, CCDchg_EndV={r["CCDchg_EndVoltage_mV"]}, Rest_BeginV={r["Rest_BeginVoltage_mV"]}");
                            Assert.False(r.IsDBNull(0), "Position should not be null");
                            Assert.False(r.IsDBNull(1), "Channel should not be null");
                            Assert.False(r.IsDBNull(2), "CCDchg_Capacity_mAh should not be null");
                        }
                        Assert.True(rowCount > 0, "No data rows found in SortingDataImportExcel_V2");
                    }
                }
            }
        }

        [Fact]
        public async Task Test_Audit_Every_Single_Cell_2_21_13_14_11()
        {
            string filePath = @"D:\ExcelData\2026-08-21\2#21_13.14.11\2#21_13.14.11.xlsx";
            if (!File.Exists(filePath)) return;

            var config = new AppConfig();
            config.Load();
            var excelService = new ExcelService(msg => _output.WriteLine(msg));

            var dt = excelService.ReadExcelFile(filePath);
            Assert.NotNull(dt);
            Assert.Equal(320, dt.Rows.Count);

            using (var conn = new SqlConnection(config.ConnectionString))
            {
                await conn.OpenAsync();
                
                // Read all DB rows into dictionary by Position + Channel
                var dbDict = new Dictionary<string, Dictionary<string, object?>>();
                using (var cmd = new SqlCommand("SELECT * FROM SortingDataImportExcel_V2 WHERE FilePath = @path", conn))
                {
                    cmd.Parameters.AddWithValue("@path", filePath);
                    using (var r = await cmd.ExecuteReaderAsync())
                    {
                        while (await r.ReadAsync())
                        {
                            string pos = r["Position"]?.ToString() ?? "";
                            string ch = r["Channel"]?.ToString() ?? "";
                            string key = $"{pos}_{ch}";
                            
                            var rowDict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                            for (int i = 0; i < r.FieldCount; i++)
                            {
                                rowDict[r.GetName(i)] = r.IsDBNull(i) ? null : r.GetValue(i);
                            }
                            dbDict[key] = rowDict;
                        }
                    }
                }

                Assert.Equal(320, dbDict.Count);

                int totalCheckedCells = 0;
                int matchedNonNullCells = 0;
                int mismatches = 0;

                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    DataRow excelRow = dt.Rows[r];
                    string pos = (dt.Columns.Contains("Position") ? excelRow["Position"] : (dt.Columns.Contains("位置") ? excelRow["位置"] : excelRow[1]))?.ToString()?.Trim() ?? "";
                    string ch = (dt.Columns.Contains("Channel") ? excelRow["Channel"] : (dt.Columns.Contains("通道") ? excelRow["通道"] : excelRow[0]))?.ToString()?.Trim() ?? "";
                    string rowKey = $"{pos}_{ch}";

                    Assert.True(dbDict.ContainsKey(rowKey), $"Row {r} with key {rowKey} not found in DB!");
                    var dbRow = dbDict[rowKey];

                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        string colName = dt.Columns[c].ColumnName;
                        string searchKey = DatabaseService.GetSearchKey(colName);
                        if (DatabaseService.AliasToSqlColumnMap.TryGetValue(searchKey, out string? sqlCol) && !string.IsNullOrEmpty(sqlCol))
                        {
                            totalCheckedCells++;
                            object excelVal = excelRow[c];
                            string excelStr = excelVal?.ToString()?.Trim() ?? "";
                            if (excelStr == "---") excelStr = "";

                            object? dbVal = dbRow.ContainsKey(sqlCol) ? dbRow[sqlCol] : null;
                            string dbStr = dbVal?.ToString()?.Trim() ?? "";

                            if (!string.IsNullOrEmpty(excelStr))
                            {
                                if (string.IsNullOrEmpty(dbStr))
                                {
                                    _output.WriteLine($"MISMATCH at row {rowKey}, col {colName} -> SQL {sqlCol}: Excel='{excelStr}' but DB is NULL/empty");
                                    mismatches++;
                                }
                                else
                                {
                                    matchedNonNullCells++;
                                }
                            }
                        }
                    }
                }

                _output.WriteLine($"\n=== 100% CELL AUDIT RESULT ===");
                _output.WriteLine($"Total Excel Rows Checked: {dt.Rows.Count}");
                _output.WriteLine($"Total DB Rows Checked: {dbDict.Count}");
                _output.WriteLine($"Total Column Cells Checked: {totalCheckedCells}");
                _output.WriteLine($"Total Non-Null Cells Matched in DB: {matchedNonNullCells}");
                _output.WriteLine($"Total Mismatches / Missing Values: {mismatches}");

                Assert.Equal(0, mismatches);
                Assert.True(matchedNonNullCells > 0);
            }
        }
    }
}
