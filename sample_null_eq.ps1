$connStr = 'Server=dbserver.hycap.co.kr,5398;Database=SmartFactoryV2;User Id=vinaadmin;Password=vina1234%6&8;TrustServerCertificate=True'
$conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
$conn.Open()
$cmd = $conn.CreateCommand()
$cmd.CommandText = "SELECT DISTINCT FilePath FROM SortingDataImportExcel WITH(NOLOCK) WHERE Barcode IS NULL"
$r = $cmd.ExecuteReader()
$paths = @()
while ($r.Read()) {
    $paths += $r.GetValue(0).ToString()
}
$conn.Close()

$existingMapped = @()
foreach ($p in $paths) {
    # Thay thế D:\net8.0-windows7.0\ bằng D:\net8.0-windows7.0 (1)\net8.0-windows7.0\
    $mappedPath = $p.Replace("D:\net8.0-windows7.0\", "D:\net8.0-windows7.0 (1)\net8.0-windows7.0\")
    if ([System.IO.File]::Exists($mappedPath)) {
        $existingMapped += $mappedPath
    }
}
Write-Host "Total existing files with mapped path: $($existingMapped.Count)"
if ($existingMapped.Count -gt 0) {
    Write-Host "Sample existing file: $($existingMapped[0])"
}
