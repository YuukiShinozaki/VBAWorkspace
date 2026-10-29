[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Extract', 'DeleteDiscard')]
    [string]$Mode,
    [string]$ProjectPath,
    [string]$WorkspaceRoot
)

$ErrorActionPreference = 'Stop'

function Resolve-ProjectPath {
    if ($ProjectPath) {
        $resolved = (Resolve-Path -LiteralPath $ProjectPath).Path
        if (-not (Test-Path -LiteralPath (Join-Path $resolved 'CopilotActivityLog.md') -PathType Leaf)) {
            throw "プロジェクト『$resolved』にCopilotActivityLog.mdが見つかりません。"
        }
        return $resolved
    }

    if (-not $WorkspaceRoot) { throw 'ProjectPathまたはWorkspaceRootを指定してください。' }
    $projectsDirectory = Join-Path (Resolve-Path -LiteralPath $WorkspaceRoot).Path 'Projects'
    $projects = @(Get-ChildItem -LiteralPath $projectsDirectory -Directory | Where-Object {
        $_.Name -notlike '_*' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'CopilotActivityLog.md') -PathType Leaf)
    })
    if ($projects.Count -eq 0) { throw 'Projects内にプロジェクトの作業記録が見つかりません。' }
    if ($projects.Count -eq 1) { return $projects[0].FullName }

    Write-Host '引継ぎメモを管理するプロジェクトを選択してください。'
    for ($i = 0; $i -lt $projects.Count; $i++) {
        Write-Host ('[{0}] {1}' -f ($i + 1), $projects[$i].Name)
    }
    $choice = Read-Host 'プロジェクト番号を入力してください（それ以外の入力でキャンセル）'
    $number = 0
    if (-not [int]::TryParse($choice, [ref]$number) -or $number -lt 1 -or $number -gt $projects.Count) {
        throw 'プロジェクトが選択されませんでした。ファイルは変更していません。'
    }
    return $projects[$number - 1].FullName
}

function Read-Entries {
    param([string]$Path)
    $text = [System.IO.File]::ReadAllText($Path)
    $matches = [regex]::Matches($text, '(?m)^(?:<!-- =+ メモ開始：[^\r\n]+ =+ -->\r?\n(?:\r?\n)*)?## \[([^\]]+)\].*(?:\r?\n|$)')
    $entries = @()
    if ($matches.Count -eq 0) { return ,$entries }
    foreach ($index in 0..($matches.Count - 1)) {
        $match = $matches[$index]
        $start = $match.Index
        $end = if ($index + 1 -lt $matches.Count) { $matches[$index + 1].Index } else { $text.Length }
        $body = $text.Substring($start, $end - $start).Trim()
        $stateMatch = [regex]::Match($body, '(?im)^\s*-\s*copilot_context:\s*(true|false|discard)\s*$')
        if ($stateMatch.Success) {
            $entries += [pscustomobject]@{ Id = $match.Groups[1].Value; Body = $body; State = $stateMatch.Groups[1].Value.ToLowerInvariant() }
        }
    }
    return ,$entries
}

try {
    $project = Resolve-ProjectPath
    $activityLog = Join-Path $project 'CopilotActivityLog.md'
    $memoryFile = Join-Path $project 'CopilotMemory.md'
    $text = [System.IO.File]::ReadAllText($activityLog)
    $entries = Read-Entries -Path $activityLog

    if ($Mode -eq 'Extract') {
        $selected = @($entries | Where-Object State -eq 'true')
        $content = @('# Copilot引継ぎメモ', '', 'CopilotActivityLog.mdから、copilot_context: trueの項目だけを抽出しています。', "生成日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')", '')
        if ($selected.Count -eq 0) {
            $content += '現在、`true`に設定された項目はありません。'
        } else {
            foreach ($entry in $selected) { $content += $entry.Body; $content += '' }
        }
        [System.IO.File]::WriteAllText($memoryFile, (($content -join "`r`n").TrimEnd() + "`r`n"), [System.Text.UTF8Encoding]::new($false))
        Write-Host "選択された$($selected.Count)件を引継ぎメモに出力しました: $memoryFile"
        exit 0
    }

    $discarded = @($entries | Where-Object State -eq 'discard')
    if ($discarded.Count -eq 0) {
        Write-Host 'discardに設定された項目はありません。ファイルは変更していません。'
        exit 0
    }
    Write-Host '以下の項目をCopilotActivityLog.mdから完全に削除します。'
    foreach ($entry in $discarded) {
        $title = [regex]::Match($entry.Body, '(?m)^##\s*([^\r\n]+)').Groups[1].Value
        Write-Host ('  {0}' -f $title)
    }
    $confirmation = Read-Host "完全に削除するには DELETE $($discarded.Count) と入力してください（バックアップは作成しません）"
    if ($confirmation -cne "DELETE $($discarded.Count)") {
        Write-Host '確認文字列が一致しません。ファイルは変更していません。'
        exit 0
    }

    $removeIds = @{}; foreach ($entry in $discarded) { $removeIds[$entry.Id] = $true }
    $sectionMatches = [regex]::Matches($text, '(?m)^(?:<!-- =+ メモ開始：[^\r\n]+ =+ -->\r?\n(?:\r?\n)*)?## \[([^\]]+)\].*(?:\r?\n|$)')
    $removeRanges = @()
    for ($i = 0; $i -lt $sectionMatches.Count; $i++) {
        $sectionStart = $sectionMatches[$i].Index
        $sectionEnd = if ($i + 1 -lt $sectionMatches.Count) { $sectionMatches[$i + 1].Index } else { $text.Length }
        $sectionText = $text.Substring($sectionStart, $sectionEnd - $sectionStart)
        $state = [regex]::Match($sectionText, '(?im)^\s*-\s*copilot_context:\s*(true|false|discard)\s*$')
        if ($removeIds.ContainsKey($sectionMatches[$i].Groups[1].Value) -and $state.Success -and $state.Groups[1].Value -ieq 'discard') {
            $removeRanges += [pscustomobject]@{ Start = $sectionStart; Length = $sectionEnd - $sectionStart }
        }
    }
    for ($i = $removeRanges.Count - 1; $i -ge 0; $i--) {
        $text = $text.Remove($removeRanges[$i].Start, $removeRanges[$i].Length)
    }
    [System.IO.File]::WriteAllText($activityLog, $text, [System.Text.UTF8Encoding]::new($false))
    Write-Host "discardの項目を$($removeRanges.Count)件削除しました。バックアップは作成していません。"
} catch {
    Write-Error $_
    exit 1
}
