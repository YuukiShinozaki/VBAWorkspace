$ErrorActionPreference = 'Stop'

$addinPath = Join-Path $PSScriptRoot 'UserFormCsvExporter.xlam'
if (-not (Test-Path -LiteralPath $addinPath -PathType Leaf)) {
    throw "アドインが見つかりません: $addinPath`nREADME.mdの手順で.xlamを作成してください。"
}

Add-Type -AssemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = 'UserForm情報を取得するExcelブックを選択'
$dialog.Filter = 'Excelマクロ有効ブック (*.xlsm;*.xlsb)|*.xlsm;*.xlsb|すべてのファイル (*.*)|*.*'
$dialog.Multiselect = $false
$dialog.CheckFileExists = $true
$dialog.InitialDirectory = (Get-Location).Path

if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host '対象ブックの選択をキャンセルしました。'
    exit 0
}

$excel = $null
$addinWorkbook = $null
$targetWorkbook = $null
$stage = 'Excelの起動'

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    # Excelの自動化セキュリティを、信頼済みのアドイン読込後に対象ブック向けへ切り替える。
    $excel.AutomationSecurity = 1
    $stage = 'アドインを開く'
    $addinWorkbook = $excel.Workbooks.Open($addinPath, 0, $true)
    $addinWorkbook.IsAddin = $true
    $excel.AutomationSecurity = 3

    $stage = '対象ブックを読み取り専用で開く'
    $targetWorkbook = $excel.Workbooks.Open($dialog.FileName, 0, $true)

    $stage = '対象ブックのVBAプロジェクトへのアクセス確認'
    try {
        $componentCount = $targetWorkbook.VBProject.VBComponents.Count
    }
    catch {
        throw "対象ブックのVBAプロジェクトを参照できません。Trust Centerの「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が無効、またはVBAプロジェクトが保護されている可能性があります。詳細: $($_.Exception.Message)"
    }

    $stage = 'UserForm CSV出力マクロの実行'
    $macroName = "'$($addinWorkbook.Name)'!Module1.ExportUserFormCsvFilesForPath"
    $formCount = [int]$excel.Run($macroName, $dialog.FileName)

    if ($formCount -eq 0) {
        Write-Host '選択したブックにUserFormはありません。'
    }
    else {
        Write-Host "UserForm情報を $formCount 件出力しました。"
        Write-Host "出力先: $(Split-Path -Parent $dialog.FileName)"
    }
}
catch {
    Write-Host "UserForm情報の出力に失敗しました。処理段階: $stage" -ForegroundColor Red
    Write-Host "詳細: $($_.Exception.ToString())"
    if ($_.Exception -is [System.Runtime.InteropServices.COMException]) {
        Write-Host ('HRESULT: 0x{0:X8}' -f $_.Exception.HResult)
    }
    Write-Host '上記の処理段階と詳細を確認してください。' -ForegroundColor Yellow
    exit 1
}
finally {
    if ($null -ne $targetWorkbook) {
        try { $targetWorkbook.Close($false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($targetWorkbook) } catch { }
    }
    if ($null -ne $addinWorkbook) {
        try { $addinWorkbook.Close($false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($addinWorkbook) } catch { }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
