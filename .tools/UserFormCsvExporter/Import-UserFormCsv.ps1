$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms

function Get-CsvValue {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][string]$Column
    )

    $value = $Row.$Column
    if ($null -eq $value) { return '' }
    return [string]$value
}

function Convert-CsvNumber {
    param(
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Column,
        [Parameter(Mandatory = $true)][int]$RowNumber,
        [switch]$Integer
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }

    $styles = [Globalization.NumberStyles]::Float
    if ($Integer) { $styles = [Globalization.NumberStyles]::Integer }

    $number = 0.0
    if (-not [double]::TryParse(
        $Text,
        $styles,
        [Globalization.CultureInfo]::CurrentCulture,
        [ref]$number
    )) {
        throw "CSVの${Column}列、${RowNumber}行目が数値ではありません: '$Text'"
    }
    if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) {
        throw "CSVの${Column}列、${RowNumber}行目に有限でない数値は指定できません。"
    }
    if ($Integer -and ($number -lt 0 -or $number -ne [math]::Truncate($number))) {
        throw "CSVの${Column}列、${RowNumber}行目は0以上の整数を指定してください。"
    }
    return $number
}

function Add-DesignerControls {
    param(
        [Parameter(Mandatory = $true)]$Controls,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$ParentName
    )

    foreach ($control in $Controls) {
        $controlName = [string]$control.Name
        $key = $ParentName + '|' + $controlName
        if ($script:controlIndex.ContainsKey($key)) {
            $script:ambiguousControlKeys[$key] = $true
        }
        else {
            $script:controlIndex[$key] = $control
        }

        $pages = $null
        try { $pages = $control.Pages } catch { }
        if ($null -ne $pages) {
            foreach ($page in $pages) {
                $pageName = [string]$page.Name
                $pageKey = $controlName + '|' + $pageName
                if ($script:controlIndex.ContainsKey($pageKey)) {
                    $script:ambiguousControlKeys[$pageKey] = $true
                }
                else {
                    $script:controlIndex[$pageKey] = $page
                }
                Add-DesignerControls -Controls $page.Controls -ParentName $pageName
            }
            continue
        }

        $children = $null
        try { $children = $control.Controls } catch { }
        if ($null -ne $children) {
            try {
                if ($children.Count -gt 0) {
                    Add-DesignerControls -Controls $children -ParentName $controlName
                }
            }
            catch { }
        }
    }
}

function Get-ValidatedActions {
    param(
        [Parameter(Mandatory = $true)]$Rows,
        [Parameter(Mandatory = $true)][string]$ExpectedFormName
    )

    $requiredColumns = @(
        'FormName', 'ParentControl', 'ControlName', 'ControlType',
        'Caption', 'Top', 'Left', 'Width', 'Height', 'TabIndex'
    )
    $headers = @($Rows[0].PSObject.Properties.Name)
    foreach ($column in $requiredColumns) {
        if ($headers -notcontains $column) {
            throw "CSVに必要な列 '$column' がありません。エクスポートしたCSVを編集して使用してください。"
        }
    }

    $actions = New-Object System.Collections.ArrayList
    $seen = @{}
    $rowNumber = 1
    foreach ($row in $Rows) {
        $rowNumber++
        $formName = Get-CsvValue -Row $row -Column 'FormName'
        $parentName = Get-CsvValue -Row $row -Column 'ParentControl'
        $controlName = Get-CsvValue -Row $row -Column 'ControlName'
        if ($formName -cne $ExpectedFormName) {
            throw "CSVのフォーム名 '$formName' が選択したフォーム '$ExpectedFormName' と一致しません。"
        }
        if ([string]::IsNullOrWhiteSpace($controlName)) {
            throw "CSVの${rowNumber}行目にControlNameがありません。"
        }

        $key = $parentName + '|' + $controlName
        if ($seen.ContainsKey($key)) {
            throw "CSVの${rowNumber}行目に同じ親・コントロール名の重複があります: $key"
        }
        $seen[$key] = $true
        if ($script:ambiguousControlKeys.ContainsKey($key)) {
            throw "Excelのフォーム内で親 '$parentName' の下に '$controlName' が複数あり、特定できません。変更は適用していません。"
        }
        if (-not $script:controlIndex.ContainsKey($key)) {
            throw "Excelに親 '$parentName' の下のコントロール '$controlName' がありません。変更は適用していません。"
        }

        $caption = Get-CsvValue -Row $row -Column 'Caption'
        $controlType = Get-CsvValue -Row $row -Column 'ControlType'
        $captionControlTypes = @(
            'Label', 'CommandButton', 'CheckBox', 'OptionButton',
            'ToggleButton', 'Frame', 'Page'
        )
        if ($captionControlTypes -notcontains $controlType -and -not [string]::IsNullOrEmpty($caption)) {
            throw "CSVの${rowNumber}行目のコントロール '$controlName' ($controlType) はCaptionを持たない種類として扱われます。Captionを空欄にしてください。"
        }
        $topText = Get-CsvValue -Row $row -Column 'Top'
        $leftText = Get-CsvValue -Row $row -Column 'Left'
        $widthText = Get-CsvValue -Row $row -Column 'Width'
        $heightText = Get-CsvValue -Row $row -Column 'Height'
        $tabIndexText = Get-CsvValue -Row $row -Column 'TabIndex'

        $action = @{
            Control = $script:controlIndex[$key]
            ControlName = $controlName
            ParentName = $parentName
            UpdateCaption = $captionControlTypes -contains $controlType
            Caption = $caption
            Top = Convert-CsvNumber -Text $topText -Column 'Top' -RowNumber $rowNumber
            Left = Convert-CsvNumber -Text $leftText -Column 'Left' -RowNumber $rowNumber
            Width = Convert-CsvNumber -Text $widthText -Column 'Width' -RowNumber $rowNumber
            Height = Convert-CsvNumber -Text $heightText -Column 'Height' -RowNumber $rowNumber
            TabIndex = Convert-CsvNumber -Text $tabIndexText -Column 'TabIndex' -RowNumber $rowNumber -Integer
        }
        [void]$actions.Add($action)
    }

    return ,$actions
}

$workbookDialog = New-Object System.Windows.Forms.OpenFileDialog
$workbookDialog.Title = 'UserFormを変更するExcelブックを選択'
$workbookDialog.Filter = 'Excelマクロ有効ブック (*.xlsm;*.xlsb)|*.xlsm;*.xlsb'
$workbookDialog.Multiselect = $false
$workbookDialog.CheckFileExists = $true
$workbookDialog.InitialDirectory = (Get-Location).Path

if ($workbookDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host '対象ブックの選択をキャンセルしました。'
    exit 0
}
$workbookPath = $workbookDialog.FileName

$csvDialog = New-Object System.Windows.Forms.OpenFileDialog
$csvDialog.Title = '変更内容を記入したUserForm CSVを選択'
$csvDialog.Filter = 'CSVファイル (*.csv)|*.csv'
$csvDialog.Multiselect = $false
$csvDialog.CheckFileExists = $true
$csvDialog.InitialDirectory = Split-Path -Parent $workbookPath

if ($csvDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host 'CSVの選択をキャンセルしました。'
    exit 0
}
$csvPath = $csvDialog.FileName

$excel = $null
$targetWorkbook = $null
$stage = 'CSVの読み込み'
$saveCompleted = $false
try {
    $rows = @(Import-Csv -LiteralPath $csvPath -Encoding UTF8)
    if ($rows.Count -eq 0) { throw 'CSVに適用できる行がありません。' }

    $formNames = @($rows | ForEach-Object { [string]$_.FormName } | Select-Object -Unique)
    if ($formNames.Count -ne 1 -or [string]::IsNullOrWhiteSpace($formNames[0])) {
        throw 'CSVには1つのUserFormのデータだけを含めてください。'
    }
    $formName = $formNames[0]

    $stage = 'Excelの起動'
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3

    $stage = '対象ブックを編集可能で開く'
    $targetWorkbook = $excel.Workbooks.Open($workbookPath, 0, $false)
    if ($targetWorkbook.ReadOnly) {
        throw '対象ブックが読み取り専用で開かれました。Excelで開いている同じブックを閉じてから再実行してください。'
    }

    $stage = 'UserFormとコントロールの確認'
    try {
        $component = $targetWorkbook.VBProject.VBComponents.Item($formName)
        $rootControls = $component.Designer.Controls
    }
    catch {
        throw "UserForm '$formName' を参照できません。Trust Centerの「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」設定、VBAProjectの保護状態、CSVのフォーム名を確認してください。詳細: $($_.Exception.Message)"
    }

    $script:controlIndex = @{}
    $script:ambiguousControlKeys = @{}
    Add-DesignerControls -Controls $rootControls -ParentName ''
    $actions = Get-ValidatedActions -Rows $rows -ExpectedFormName $formName

    $preview = @($actions | Select-Object -First 8 | ForEach-Object {
        "{0} / {1}: Caption='{2}', Top={3}, Left={4}, Width={5}, Height={6}, TabIndex={7}" -f `
            $_.ParentName, $_.ControlName, $_.Caption, $_.Top, $_.Left, $_.Width, $_.Height, $_.TabIndex
    }) -join "`r`n"
    if ($actions.Count -gt 8) { $preview += "`r`n...ほか $($actions.Count - 8) 件" }

    $message = "UserForm '$formName' の既存コントロール $($actions.Count) 件をCSVの値で更新します。`r`n`r`n$preview`r`n`r`n続行前にブックのバックアップを作成し、適用後にブックを保存します。実行しますか？"
    $choice = [System.Windows.Forms.MessageBox]::Show(
        $message,
        'UserForm CSVの反映確認',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host '反映をキャンセルしました。'
        exit 0
    }

    $stage = 'ブックのバックアップ'
    $directory = Split-Path -Parent $workbookPath
    $baseName = [IO.Path]::GetFileNameWithoutExtension($workbookPath)
    $extension = [IO.Path]::GetExtension($workbookPath)
    $backupPath = Join-Path $directory ($baseName + '.backup-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + $extension)
    Copy-Item -LiteralPath $workbookPath -Destination $backupPath

    $stage = 'UserFormの見た目を更新'
    foreach ($action in $actions) {
        $control = $action.Control
        if ($action.UpdateCaption) { $control.Caption = $action.Caption }
        if ($null -ne $action.Top) { $control.Top = [single]$action.Top }
        if ($null -ne $action.Left) { $control.Left = [single]$action.Left }
        if ($null -ne $action.Width) { $control.Width = [single]$action.Width }
        if ($null -ne $action.Height) { $control.Height = [single]$action.Height }
        if ($null -ne $action.TabIndex) { $control.TabIndex = [int]$action.TabIndex }
    }

    $stage = 'ブックの保存'
    $targetWorkbook.Save()
    $saveCompleted = $true
    Write-Host "CSVの内容をUserForm '$formName' に反映しました。"
    Write-Host "バックアップ: $backupPath"
}
catch {
    Write-Host "UserForm CSVの反映に失敗しました。処理段階: $stage" -ForegroundColor Red
    Write-Host "詳細: $($_.Exception.ToString())"
    if ($_.Exception -is [System.Runtime.InteropServices.COMException]) {
        Write-Host ('HRESULT: 0x{0:X8}' -f $_.Exception.HResult)
    }
    if ($null -ne $backupPath -and -not $saveCompleted) {
        Write-Host "保存前のバックアップは残っています: $backupPath" -ForegroundColor Yellow
    }
    exit 1
}
finally {
    if ($null -ne $targetWorkbook) {
        try { $targetWorkbook.Close($false) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($targetWorkbook) } catch { }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
