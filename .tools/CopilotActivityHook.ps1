$ErrorActionPreference = 'Stop'

function Write-HookOutput {
    [Console]::Out.WriteLine('{}')
}

function Find-ProjectRoot {
    param([string]$WorkingDirectory)

    if ([string]::IsNullOrWhiteSpace($WorkingDirectory)) { return $null }
    $current = [System.IO.Path]::GetFullPath($WorkingDirectory)
    while ($current) {
        $activityLog = Join-Path $current 'CopilotActivityLog.md'
        if ((Test-Path -LiteralPath $activityLog -PathType Leaf) -and
            (Split-Path -Leaf $current) -ne '_Template_Excel') {
            return $current
        }
        $parent = Split-Path -Parent $current
        if (-not $parent -or $parent -eq $current) { break }
        $current = $parent
    }

    $projectsDirectory = Join-Path $WorkingDirectory 'Projects'
    if (-not (Test-Path -LiteralPath $projectsDirectory -PathType Container)) {
        $ancestor = [System.IO.DirectoryInfo]::new($WorkingDirectory)
        while ($ancestor -and -not (Test-Path -LiteralPath (Join-Path $ancestor.FullName 'Projects') -PathType Container)) {
            $ancestor = $ancestor.Parent
        }
        if ($ancestor) { $projectsDirectory = Join-Path $ancestor.FullName 'Projects' }
    }
    if (Test-Path -LiteralPath $projectsDirectory -PathType Container) {
        $candidates = @(Get-ChildItem -LiteralPath $projectsDirectory -Directory | Where-Object {
            $_.Name -notlike '_*' -and (Test-Path -LiteralPath (Join-Path $_.FullName 'CopilotActivityLog.md') -PathType Leaf)
        })
        if ($candidates.Count -eq 1) { return $candidates[0].FullName }
    }
    return $null
}

try {
    $inputJson = [Console]::In.ReadToEnd()
    if ([string]::IsNullOrWhiteSpace($inputJson)) { Write-HookOutput; exit 0 }
    $event = $inputJson | ConvertFrom-Json
    $prompt = [string]$event.prompt
    if ([string]::IsNullOrWhiteSpace($prompt)) { Write-HookOutput; exit 0 }

    $actionPattern = '(?i)(作成して|作成してください|作って|実装して|実装してください|コード.{0,12}(作成|追加|変更)|修正して|修正してください|直して|変更して|変更してください|追加して|追加してください|編集して|編集してください|更新して|更新してください|削除して|削除してください|置き換えて|書き換えて|対応して|対応してください|反映して|反映してください|構築して|作成をお願い|実装をお願い|修正をお願い|add\s+(?:a|the|new)?\s*(?:feature|function|file|test|task)|create\s+(?:a|the|new)?\s*(?:feature|function|file|script|task)|implement\s+(?:a|the|this|the requested)?|fix\s+(?:the|this|a)?\s*(?:bug|issue|error|code)|update\s+(?:the|this|a)?\s*(?:file|code|configuration|config|document)|modify\s+(?:the|this|a)?\s*(?:file|code|configuration|config)|change\s+(?:the|this|a)?\s*(?:file|code|behavior)|please\s+(?:add|create|implement|fix|update|modify|change|edit|remove))'
    $discussionPattern = '(?i)(どう思いますか|意見を聞|相談(?:なん|した|です)|案として|可能でしょうか|できますか|した方がよい|したほうがいい|検討したい|考えを聞|what do you think|should we|could we|is it possible|would it be better|let us discuss)'
    $negationPattern = '(?i)(作成しないで|作成は不要|実装しないで|実装は不要|変更しないで|修正しないで|削除しないで|まだ実行しない|まだ作らない|do not (?:create|implement|change|modify|fix|update)|don.t (?:create|implement|change|modify|fix|update)|not yet|no changes? yet)'
    if ($prompt -notmatch $actionPattern -or $prompt -match $negationPattern) { Write-HookOutput; exit 0 }
    if ($prompt -match $discussionPattern -and $prompt -notmatch '(?i)(実装して|作成して|修正して|変更して|追加して|更新して|implement|create|fix|update|add|modify|please)') {
        Write-HookOutput; exit 0
    }

    $projectRoot = Find-ProjectRoot -WorkingDirectory ([string]$event.cwd)
    if (-not $projectRoot) { Write-HookOutput; exit 0 }
    $activityLog = Join-Path $projectRoot 'CopilotActivityLog.md'
    $safePrompt = $prompt -replace '(?i)(password|passwd|api[_-]?key|client[_-]?secret|access[_-]?token|authorization)(\s*[:=]\s*)\S+', '$1$2[REDACTED]'
    $safePrompt = $safePrompt -replace '(?i)Bearer\s+[A-Za-z0-9._~+/-]+=*', 'Bearer [REDACTED]'
    $safePrompt = $safePrompt -replace '(?im)(-----BEGIN [^-]+PRIVATE KEY-----).*?(-----END [^-]+PRIVATE KEY-----)', '$1[REDACTED]$2'
    $safePrompt = $safePrompt.Trim()
    $id = 'CA-{0}-{1}' -f (Get-Date -Format 'yyyyMMdd'), ([Guid]::NewGuid().ToString('N').Substring(0, 8).ToUpperInvariant())
    $quotedPrompt = (($safePrompt -split "`r?`n") | ForEach-Object { '> ' + $_ }) -join "`r`n"
    $entry = "`r`n`r`n<!-- ==================== メモ開始：$id ==================== -->`r`n`r`n## [$id] 作業依頼`r`n`r`n- copilot_context: false`r`n- 記録日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')`r`n- 記録元: プロンプト送信フック（userPromptSubmitted）`r`n`r`n### 依頼内容`r`n$quotedPrompt`r`n`r`n<!-- ==================== メモ終了：$id ==================== -->`r`n"
    [System.IO.File]::AppendAllText($activityLog, $entry, [System.Text.UTF8Encoding]::new($false))
} catch {
    [Console]::Error.WriteLine("Copilot作業記録の自動追記をスキップしました: $($_.Exception.Message)")
}

Write-HookOutput
