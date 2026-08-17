-- ============================================================================
-- BACKUP STORED PROCEDURE ORIGINAL (Trước khi sửa)
-- Stored Procedure: [dbo].[sp_GetSortingDataProgram]
-- Backup Date: 2026-08-17
-- Database: SmartFactoryV2
-- ============================================================================

CREATE PROCEDURE [dbo].[sp_GetSortingDataProgram]
    @pProcessUserID VARCHAR(20) = NULL,
    @pProcessLanguage VARCHAR(20) = NULL,
    @pFromDate DATETIME,
    @pToDate   DATETIME,
    @pBarcode NVARCHAR(50) = NULL,
    @pEquipmentNumber NVARCHAR(50) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @FromDate VARCHAR(19) = CONVERT(VARCHAR(10), @pFromDate, 120)
    DECLARE @ToDate   VARCHAR(19) = CONVERT(VARCHAR(10), DATEADD(DAY, 1, @pToDate), 120) 
    DECLARE @Barcode  VARCHAR(50) = CASE WHEN ISNULL(@pBarcode, '') = '' THEN '*' ELSE @pBarcode END
    DECLARE @EquipmentNumber VARCHAR(50) = CASE WHEN ISNULL(@pEquipmentNumber, '') = '' THEN '*' ELSE @pEquipmentNumber END

    -- CTE tra cứu nguyên vật liệu khi có mã Barcode
    ;WITH RawMaterial AS (
        SELECT 
            RM.Barcode,
            SUBSTRING(MAX(CASE WHEN RM.ProductGroupCode = 'ElectrodeM' THEN RM.RawMaterialBarcode END), 1, 14) AS ElectrodeM_Barcode_Short,
            SUBSTRING(MAX(CASE WHEN RM.ProductGroupCode = 'ElectrodeP' THEN RM.RawMaterialBarcode END), 1, 14) AS ElectrodeP_Barcode_Short
        FROM STB_RawMaterialInputHist RM WITH(NOLOCK)
        WHERE RM.CreateUserID = 'vvtworker_bg' 
          AND RM.ProductGroupCode IN ('ElectrodeM','ElectrodeP')
          AND (@Barcode = '*' OR RM.Barcode = @Barcode)
        GROUP BY RM.Barcode
    ),
    MappedMaterials AS (
        SELECT 
            RM.*,
            SetM.MaterialCode AS ElectrodeM_Code,
            SetP.MaterialCode AS ElectrodeP_Code
        FROM RawMaterial RM
        LEFT JOIN STB_SetInfo SetM WITH(NOLOCK) ON RM.ElectrodeM_Barcode_Short = SetM.Barcode
        LEFT JOIN STB_SetInfo SetP WITH(NOLOCK) ON RM.ElectrodeP_Barcode_Short = SetP.Barcode
    ),
    CombinedSortingData AS (
        -- 1. Dữ liệu File Excel Cũ (Ver 1)
        SELECT 
            EquipmentNumber,
            SorterNum,
            DATEADD(hour, 2, StartTime) AS StartTime,
            ISNULL(NULLIF(WorkflowCode, ''), '35105') AS WorkflowCode,
            Barcode,
            Slot,
            Position,
            Channel,
            TRY_CAST(Capacity_mAh AS FLOAT) AS Capacity_mAh,
            TRY_CAST(Capacitance_F AS FLOAT) AS Capacitance_F,
            TRY_CAST(BeginVoltageSD_mV AS FLOAT) AS BeginVoltageSD_mV,
            TRY_CAST(ChargeEndCurrent_mA AS FLOAT) AS ChargeEndCurrent_mA,
            TRY_CAST(EndVoltage_mV AS FLOAT) AS EndVoltage_mV,
            TRY_CAST(EndCurrent_mA AS FLOAT) AS EndCurrent_mA,
            TRY_CAST(DischargeVoltage1_mV AS FLOAT) AS DischargeVoltage1_mV,
            CAST(DischargeVoltage1_Time AS NVARCHAR(50)) AS DischargeVoltage1_Time,
            TRY_CAST(DischargeVoltage2_mV AS FLOAT) AS DischargeVoltage2_mV,
            CAST(DischargeVoltage2_Time AS NVARCHAR(50)) AS DischargeVoltage2_Time,
            TRY_CAST(DischargeBeginVoltage_mV AS FLOAT) AS DischargeBeginVoltage_mV,
            TRY_CAST(DischargeBeginCurrent_mA AS FLOAT) AS DischargeBeginCurrent_mA,
            TRY_CAST(ESR_mOhm AS FLOAT) AS ESR_mOhm,
            TRY_CAST(OCV_mV AS FLOAT) AS OCV_mV,
            ESRTime,
            NGInfo,
            DATEADD(hour, 2, EndTime) AS EndTime,
            FilePath,
            ImportDate
        FROM SortingDataImportExcel WITH(NOLOCK)
        WHERE (ImportDate >= @FromDate AND ImportDate <= @ToDate)
          AND (@Barcode = '*' OR @Barcode = '' OR Barcode = @Barcode)
          AND (@EquipmentNumber = '*' OR @EquipmentNumber = '' OR EquipmentNumber LIKE '%' + @EquipmentNumber + '%')

        UNION ALL

        -- 2. Dữ liệu File Excel Mới (Ver 2)
        SELECT 
            EquipmentNumber,
            ISNULL(NULLIF(TrayID, ''), 
                CASE WHEN CHARINDEX('#', Position) > 0 AND CHARINDEX('-', Position) > CHARINDEX('#', Position)
                     THEN SUBSTRING(Position, CHARINDEX('#', Position) + 1, CHARINDEX('-', Position) - CHARINDEX('#', Position) - 1)
                     ELSE TrayID END) AS SorterNum,
            CCCVChg_BeginTime AS StartTime,
            '35105' AS WorkflowCode,
            Barcode,
            CASE WHEN CHARINDEX('-', Position) > 0 
                 THEN SUBSTRING(Position, CHARINDEX('-', Position) + 1, 10) 
                 ELSE '' END AS Slot,
            Position,
            Channel,
            CCDchg_Capacity_mAh AS Capacity_mAh,
            CCDchg_Capacitance_F AS Capacitance_F,
            CCCVChg_BeginVoltage_mV AS BeginVoltageSD_mV,
            CCCVChg_EndCurrent_mA AS ChargeEndCurrent_mA,
            CCCVChg_EndVoltage_mV AS EndVoltage_mV,
            CCCVChg_EndCurrent_mA AS EndCurrent_mA,
            CCDchg_BeginVoltage_mV AS DischargeVoltage1_mV,
            CCDchg_WorkstepTime AS DischargeVoltage1_Time,
            CCDchg_EndVoltage_mV AS DischargeVoltage2_mV,
            Rest_WorkstepTime AS DischargeVoltage2_Time,
            CCDchg_BeginVoltage_mV AS DischargeBeginVoltage_mV,
            CCDchg_BeginCurrent_mA AS DischargeBeginCurrent_mA,
            NULL AS ESR_mOhm,
            NULL AS OCV_mV,
            NULL AS ESRTime,
            CCDchg_StopReason AS NGInfo,
            Rest_EndTime AS EndTime,
            FilePath,
            ImportDate
        FROM SortingDataImportExcel_V2 WITH(NOLOCK)
        WHERE (ImportDate >= @FromDate AND ImportDate <= @ToDate)
          AND (@Barcode = '*' OR @Barcode = '' OR Barcode = @Barcode)
          AND (@EquipmentNumber = '*' OR @EquipmentNumber = '' OR EquipmentNumber LIKE '%' + @EquipmentNumber + '%')
    )

    -- 3. Trả về ĐÚNG 100% TÊN CỘT CHUẨN CỦA MÀN W788 NATIVE
    SELECT 
        ISNULL(T2.InputLineCode, 'Line ' + REPLACE(T1.EquipmentNumber, '#', '')) AS InputLineCode,
        'Sorting ' + REPLACE(T1.EquipmentNumber, '#', '') AS EquipmentNumber,
        T1.SorterNum,
        T1.StartTime,
        T1.WorkflowCode,
        T2.LotNumber AS LotNo,
        T1.Barcode,
        T1.Slot,
        T1.Position,
        T1.Channel,
        T1.Capacity_mAh,
        T1.Capacitance_F,
        T1.BeginVoltageSD_mV,
        T1.ChargeEndCurrent_mA,
        T1.EndVoltage_mV,
        T1.EndCurrent_mA,
        T1.DischargeVoltage1_mV,
        T1.DischargeVoltage1_Time,
        T1.DischargeVoltage2_mV,
        T1.DischargeVoltage2_Time,
        T1.DischargeBeginVoltage_mV,
        T1.DischargeBeginCurrent_mA,
        T1.ESR_mOhm,
        T1.OCV_mV,
        T1.ESRTime,
        T1.NGInfo,
        T1.EndTime,
        T1.FilePath,
        T10.ElectrodeM_Barcode_Short AS ElectrodeM_Code,
        T3.MaterialName AS ElectrodeM_Name,
        T10.ElectrodeP_Barcode_Short AS ElectrodeP_Code,
        T4.MaterialName AS ElectrodeP_Name
    FROM CombinedSortingData T1
    LEFT JOIN STB_SetInfo T2 WITH(NOLOCK) ON ISNULL(T1.Barcode, '') <> '' AND T1.Barcode = T2.Barcode
    LEFT JOIN MappedMaterials T10 ON ISNULL(T1.Barcode, '') <> '' AND T1.Barcode = T10.Barcode 
    LEFT JOIN STB_MaterialMaster T3 WITH(NOLOCK) ON T10.ElectrodeM_Code = T3.MaterialCode
    LEFT JOIN STB_MaterialMaster T4 WITH(NOLOCK) ON T10.ElectrodeP_Code = T4.MaterialCode
    ORDER BY T1.ImportDate DESC;
END
