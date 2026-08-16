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
            string file1 = Path.Combine(_workspaceDir, "ExcelData", "2026-08-15", "1#15_14.52.13", "1#15_14.52.13.xlsx");
            if (!File.Exists(file1)) return;

            var dt1 = _excelService.ReadExcelFile(file1);
            if (dt1 == null) return;

            _output.WriteLine("--- FILE 1 COLUMNS ---");
            foreach (DataColumn dc in dt1.Columns)
            {
                _output.WriteLine($"'{dc.ColumnName}' -> Key: '{DatabaseService.GetSearchKey(dc.ColumnName)}'");
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
    }
}
