using System;
using Xunit;
using ImportData.Services;

namespace ImportData.Tests
{
    public class DatabaseServiceTests
    {
        [Fact]
        public void Test_GetSearchKey_Normalization()
        {
            // Test zero-width space
            Assert.Equal("equipmentnumber", DatabaseService.GetSearchKey("Equipment Number\u200B"));

            // Test newline and parentheses
            Assert.Equal("chargeendcurrentma", DatabaseService.GetSearchKey("Charge\nEndCurrent(mA)"));

            // Test spaces, colons, underscores
            Assert.Equal("dischargevoltage1timemmss", DatabaseService.GetSearchKey("DischargeVoltage1_Time (mm：ss)"));

            // Test BOM character
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
        [InlineData(102, false)]  // Syntax error (Not transient)
        [InlineData(547, false)]  // FK Constraint violation (Not transient)
        [InlineData(2627, false)] // Unique constraint violation (Not transient)
        public void Test_IsTransientErrorNumber(int errorCode, bool expectedIsTransient)
        {
            bool actual = DatabaseService.IsTransientErrorNumber(errorCode);
            Assert.Equal(expectedIsTransient, actual);
        }
    }
}
