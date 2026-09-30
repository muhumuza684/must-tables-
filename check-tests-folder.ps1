<#
 Data Lake Tables - check the tests folder (read-only, changes nothing)

 Jest is finding the tests folder but zero *.test.ts files inside it.
 This just looks and reports - no files are touched.

 Save in the project folder and run:  .\check-tests-folder.ps1
#>
$testsPath = Join-Path (Get-Location) "tests"

Write-Host "Path: $testsPath"
Write-Host ""

if (-not (Test-Path $testsPath)) {
    Write-Host "The tests folder does not exist at that path." -ForegroundColor Red
    exit
}

$all = Get-ChildItem -Path $testsPath -Force
Write-Host "Items directly inside tests\ : $($all.Count)"
foreach ($item in $all) {
    $attrs = $item.Attributes
    Write-Host "  $($item.Name)   [$attrs]"
}

Write-Host ""
$testFiles = Get-ChildItem -Path $testsPath -Filter "*.test.ts" -File -Force -ErrorAction SilentlyContinue
Write-Host "Files matching *.test.ts directly inside tests\ : $($testFiles.Count)"

Write-Host ""
Write-Host "Checking OneDrive placeholder status (cloud-only files show here if any):" -ForegroundColor Cyan
foreach ($item in $all) {
    if ($item.Attributes -match "Offline" -or $item.Attributes -match "RecallOnDataAccess") {
        Write-Host "  CLOUD-ONLY (not downloaded): $($item.Name)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "What Jest itself sees:" -ForegroundColor Cyan
npx jest --listTests
