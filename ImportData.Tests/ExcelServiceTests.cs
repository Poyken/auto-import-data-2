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
            // Thiết lập logger rỗng cho test
            _excelService = new ExcelService(msg => Console.WriteLine(msg));

            // Tìm thư mục workspace chứa 2 file Excel mẫu
            string currentDir = AppDomain.CurrentDomain.BaseDirectory;
            while (currentDir != null && !File.Exists(Path.Combine(currentDir, "20260531191549-6#A2.xlsx")))
            {
                currentDir = Path.GetDirectoryName(currentDir)!;
            }
            _workspaceDir = currentDir;
        }

        [Fact]
        public void PrintColumns()
        {
            var dt1 = _excelService.ReadExcelFile(Path.Combine(_workspaceDir, "20260531191549-6#A2.xlsx"));
            var dt2 = _excelService.ReadExcelFile(Path.Combine(_workspaceDir, "20260601000502-6#A2.xlsx"));

            _output.WriteLine("--- FILE 1 COLUMNS ---");
            foreach (DataColumn dc in dt1.Columns)
            {
                _output.WriteLine($"'{dc.ColumnName}' -> '{ExcelService.NormalizeColumnName(dc.ColumnName)}' -> Key: '{DatabaseService.GetSearchKey(dc.ColumnName)}'");
            }
            if (dt1.Rows.Count > 0)
            {
                _output.WriteLine("--- FILE 1 FIRST ROW ---");
                foreach (DataColumn dc in dt1.Columns)
                {
                    _output.WriteLine($"'{dc.ColumnName}': '{dt1.Rows[0][dc]}'");
                }
            }

            _output.WriteLine("--- FILE 2 COLUMNS ---");
            foreach (DataColumn dc in dt2.Columns)
            {
                _output.WriteLine($"'{dc.ColumnName}' -> '{ExcelService.NormalizeColumnName(dc.ColumnName)}' -> Key: '{DatabaseService.GetSearchKey(dc.ColumnName)}'");
            }
            if (dt2.Rows.Count > 0)
            {
                _output.WriteLine("--- FILE 2 FIRST ROW ---");
                foreach (DataColumn dc in dt2.Columns)
                {
                    _output.WriteLine($"'{dc.ColumnName}': '{dt2.Rows[0][dc]}'");
                }
            }
        }

        [Fact]
        public void Test_NormalizeColumnName_HiddenCharsAndNewlines()
        {
            // Test zero-width space
            string col1 = "Equipment Number\u200B";
            Assert.Equal("equipmentnumber", ExcelService.NormalizeColumnName(col1));

            // Test newline and parentheses
            string col2 = "Charge\nEndCurrent(mA)";
            Assert.Equal("chargeendcurrentma", ExcelService.NormalizeColumnName(col2));

            // Test spaces, fullwidth colon, and underscores
            string col3 = "DischargeVoltage1_Time (mm：ss)";
            Assert.Equal("dischargevoltage1timemmss", ExcelService.NormalizeColumnName(col3));

            // Test BOM character
            string col4 = "\uFEFFBarcode";
            Assert.Equal("barcode", ExcelService.NormalizeColumnName(col4));
        }

        [Fact]
        public void Test_ReadExcelFile_Format1_EquipmentNumber()
        {
            string filePath = Path.Combine(_workspaceDir, "20260531191549-6#A2.xlsx");
            Assert.True(File.Exists(filePath), $"Không tìm thấy file test format 1 tại: {filePath}");

            DataTable dt = _excelService.ReadExcelFile(filePath);
            
            Assert.NotNull(dt);
            // Kiểm tra ValidateHeaders hoạt động đúng
            Assert.True(_excelService.ValidateHeaders(dt));
            // Quy ước thông thường có 64 dòng dữ liệu
            Assert.Equal(64, dt.Rows.Count);
        }

        [Fact]
        public void Test_ReadExcelFile_Format2_DevName()
        {
            string filePath = Path.Combine(_workspaceDir, "20260601000502-6#A2.xlsx");
            Assert.True(File.Exists(filePath), $"Không tìm thấy file test format 2 tại: {filePath}");

            DataTable dt = _excelService.ReadExcelFile(filePath);

            Assert.NotNull(dt);
            // Kiểm tra ValidateHeaders hoạt động đúng
            Assert.True(_excelService.ValidateHeaders(dt));
            // Quy ước thông thường có 64 dòng dữ liệu
            Assert.Equal(64, dt.Rows.Count);
        }

        [Fact]
        public void Test_HasEsrColumns()
        {
            string file1 = Path.Combine(_workspaceDir, "20260531191549-6#A2.xlsx");
            string file2 = Path.Combine(_workspaceDir, "20260601000502-6#A2.xlsx");

            Assert.False(_excelService.HasEsrColumns(file1));
            Assert.True(_excelService.HasEsrColumns(file2));
        }

        [Fact]
        public void Test_NormalizeColumnName_ExtremeInputs()
        {
            // Null & Empty
            Assert.Equal(string.Empty, ExcelService.NormalizeColumnName(null!));
            Assert.Equal(string.Empty, ExcelService.NormalizeColumnName(""));

            // Tab characters, multiple spaces, mixed special symbols
            string input1 = "\tEquipment   \n   Number_#!%";
            Assert.Equal("equipmentnumber#!", ExcelService.NormalizeColumnName(input1));

            // Carriage return and backslash
            string input2 = "Discharge\r/Voltage\\1_Time";
            Assert.Equal("dischargevoltage1time", ExcelService.NormalizeColumnName(input2));
        }

        [Fact]
        public void Test_ValidateHeaders_Thresholds()
        {
            DataTable dt = new DataTable();
            
            // Ít hơn 5 cột
            dt.Columns.Add("Barcode");
            dt.Columns.Add("StartTime");
            Assert.False(_excelService.ValidateHeaders(dt));

            // Có hơn 5 cột nhưng không có cột nào khớp header quan trọng
            dt.Columns.Add("Unrelated1");
            dt.Columns.Add("Unrelated2");
            dt.Columns.Add("Unrelated3");
            Assert.False(_excelService.ValidateHeaders(dt));

            // Thêm các cột khớp tiêu đề quan trọng
            dt.Columns.Add("SorterNum");
            dt.Columns.Add("Slot");
            dt.Columns.Add("Position");
            dt.Columns.Add("Channel");
            dt.Columns.Add("Capacity");
            dt.Columns.Add("Capacitance");
            
            // Tổng cộng có: Barcode, StartTime, SorterNum, Slot, Position, Channel, Capacity, Capacitance (8 cột khớp)
            Assert.True(_excelService.ValidateHeaders(dt));
        }

        [Fact]
        public void Test_CleanEmptyCells_And_ReplaceDashesWithDbNull()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Col1", typeof(object));
            dt.Columns.Add("Col2", typeof(object));

            // Thêm dòng hợp lệ
            dt.Rows.Add("Value1", "Value2");
            
            // Thêm dòng trống (hoặc toàn gạch ngang)
            dt.Rows.Add("---", "---");
            dt.Rows.Add("   ", "");
            dt.Rows.Add("Value3", "---");

            // Thực thi logic làm sạch tương tự trong ExcelService.ReadExcelFile
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

            // Kết quả mong muốn:
            // - Dòng 1 ("Value1", "Value2") giữ nguyên
            // - Dòng 2 ("---", "---") bị xóa bỏ hoàn toàn
            // - Dòng 3 ("   ", "") bị xóa bỏ hoàn toàn
            // - Dòng 4 ("Value3", "---") được giữ nhưng phần tử "---" bị thay bằng DBNull.Value

            Assert.Equal(2, dt.Rows.Count);
            
            // Dòng thứ nhất
            Assert.Equal("Value1", dt.Rows[0]["Col1"]);
            Assert.Equal("Value2", dt.Rows[0]["Col2"]);

            // Dòng thứ hai (dòng 4 cũ)
            Assert.Equal("Value3", dt.Rows[1]["Col1"]);
            Assert.Equal(DBNull.Value, dt.Rows[1]["Col2"]);
        }
    }
}
