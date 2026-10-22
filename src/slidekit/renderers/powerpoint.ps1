param(
    [ValidateSet('Session', 'Status', 'Clear')][string]$Mode,
    [string]$StateDir,
    [string]$WorkDir,
    [string]$RunId,
    [double]$WaitSeconds = 0,
    [string]$Holder = '',
    [switch]$Force
)

# slidekit の PowerPoint の作業スクリプト。Windows 利用者ごとの排他で、変換を一度に一つにする。
# Session：排他を取り、標準入力のリクエスト（1 行に 1 つの JSON）に応答する。標準入力の終了は close と同じ扱い。
# Status：排他・残存記録・使用中の処理を調べる。Clear：安全を確かめてから残存記録を解除する。
# 利用者の PowerPoint は強制終了しない。開いた複製だけを閉じ、自分で起動して資料が残っていない場合だけ終了する。
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
try { [Console]::InputEncoding = $utf8 } catch { }
$sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$mutex = New-Object System.Threading.Mutex($false, "Global\Slidekit.PowerPoint.$sid")
$pending = Join-Path $StateDir 'pending.json'
$active = Join-Path $StateDir 'active.json'
$locked = $false
$abandoned = $false
$canWrite = $true
$app = $null
$existing = $false
$holderData = $null

function Send-Line([string]$text) {
    # Python 側が切り離した後（パイプが閉じた後）は書くことができない。後片付けは続ける。
    if (-not $script:canWrite) { return }
    try {
        [Console]::Out.WriteLine($text)
        [Console]::Out.Flush()
    } catch {
        $script:canWrite = $false
    }
}

function Send-Event($data) {
    Send-Line ('SLIDEKIT_EVENT:' + (ConvertTo-Json -InputObject $data -Compress -Depth 8))
}

function Send-Result($data) {
    Send-Line ('SLIDEKIT_JSON:' + (ConvertTo-Json -InputObject $data -Compress -Depth 8))
}

function Read-JsonFile([string]$path) {
    try {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
        return ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($path, [Text.Encoding]::UTF8))
    } catch {
        return $null
    }
}

function Write-JsonFile([string]$path, $data) {
    [IO.File]::WriteAllText($path, (ConvertTo-Json -InputObject $data -Compress -Depth 8), $utf8)
}

function Lock-Mutex([bool]$wait) {
    # 1 秒ずつ待つ。保持していた処理が終了させられていた場合も取得として扱う。
    $acquired = $false
    try {
        if ($wait) { $acquired = $mutex.WaitOne(1000) } else { $acquired = $mutex.WaitOne(0) }
    } catch [System.Threading.AbandonedMutexException] {
        $script:abandoned = $true
        $acquired = $true
    }
    if ($acquired) { $script:locked = $true }
    return $acquired
}

function Unlock-Mutex {
    if ($script:locked) {
        try { $mutex.ReleaseMutex() } catch { }
        $script:locked = $false
    }
}

function Format-Holder($record) {
    # 使用中の記録から「（PID …、開始 …、資料 …、実行 …）」を作る。分からない項目は省く。
    if ($null -eq $record) { return '' }
    $parts = @()
    if ($record.pid) { $parts += "PID $($record.pid)" }
    if ($record.started_at) { $parts += "開始 $($record.started_at)" }
    if ($null -ne $record.holder) {
        if ($record.holder.project) { $parts += "資料 $($record.holder.project)" }
        if ($record.holder.run) { $parts += "実行 $($record.holder.run)" }
    }
    if ($parts.Count -eq 0) { return '' }
    return '（' + ($parts -join '、') + '）'
}

function Get-RegisteredFonts {
    $fonts = @()
    foreach ($key in @(
        'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts',
        'HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    )) {
        if (-not (Test-Path -LiteralPath $key)) { continue }
        $reg = Get-Item -LiteralPath $key
        foreach ($name in $reg.GetValueNames()) {
            $value = [string]$reg.GetValue($name)
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            $value = [Environment]::ExpandEnvironmentVariables($value)
            if (-not [IO.Path]::IsPathRooted($value)) {
                $base = Join-Path $env:WINDIR 'Fonts'
                if ($key.StartsWith('HKCU:')) {
                    $base = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
                }
                $value = Join-Path $base $value
            }
            if (Test-Path -LiteralPath $value -PathType Leaf) {
                $fonts += @{name=$name; path=$value}
            }
        }
    }
    return $fonts
}

function Invoke-Doctor($id) {
    # 事前確認では COM を起動しないため、開いている資料には触れない。
    try {
        $type = [Type]::GetTypeFromProgID('PowerPoint.Application')
        if ($null -eq $type) { throw 'PowerPoint.Application が登録されていません。' }
        $key = "Registry::HKEY_CLASSES_ROOT\CLSID\{$($type.GUID)}\LocalServer32"
        $command = [string](Get-Item -LiteralPath $key).GetValue('')
        $command = [Environment]::ExpandEnvironmentVariables($command)
        if ($command -match '^\s*"([^"]+)"') { $executable = $matches[1] }
        elseif ($command -match '^\s*(.+?\.exe)(?:\s|$)') { $executable = $matches[1] }
        else { throw 'PowerPoint の実行ファイルを特定できません。' }
        if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
            throw "PowerPoint の実行ファイルがありません: $executable"
        }
        $version = [Diagnostics.FileVersionInfo]::GetVersionInfo($executable)
        return @{id=$id; ok=$true; renderer='powerpoint'; version=$version.ProductVersion;
                 executable=$executable; windows_user_sid=$sid;
                 installed_fonts=@(Get-RegisteredFonts); state_dir=$StateDir}
    } catch {
        return @{id=$id; ok=$false; renderer='powerpoint'; message=$_.Exception.Message;
                 state_dir=$StateDir; work_dir=$WorkDir; run_id=$RunId}
    }
}

function Test-Disconnected($exception) {
    # PowerPoint との接続が切れた（RPC_S_SERVER_UNAVAILABLE 0x800706BA、RPC_E_DISCONNECTED 0x80010108）。
    while ($null -ne $exception) {
        if ($exception.HResult -eq -2147023174 -or $exception.HResult -eq -2147417848) { return $true }
        $exception = $exception.InnerException
    }
    return $false
}

function New-App {
    # 既に動いている PowerPoint があれば、それに接続する（終了させない）。自分で起動したかを返す。
    # 前の処理が終了させている途中の PowerPoint があると、COM は新しい PowerPoint を起動するので、PowerPoint があったかでなく、
    # 接続の前後の処理の番号を比べて、新しい処理が増えたときだけ自分で起動したとする。
    $before = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $script:app = New-Object -ComObject PowerPoint.Application
    $after = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $script:existing = @($after | Where-Object { $before -notcontains $_ }).Count -eq 0
    return (-not $script:existing)
}

function Reset-App {
    if ($null -ne $script:app) {
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($script:app) | Out-Null } catch { }
    }
    $script:app = $null
}

function Invoke-Render($request) {
    $id = $request.id
    $inputPath = [string]$request.input
    $outputPath = [string]$request.output
    $presentation = $null
    $recorded = $false
    $started = $false
    try {
        if (Test-Path -LiteralPath $pending) {
            throw "SLIDEKIT_PENDING: 前回の処理の終了を確認してください。記録: $pending"
        }
        if (Test-Path -LiteralPath $outputPath) { throw '新しい PDF 出力先が必要です。' }
        # 呼び出しが戻らない場合に備え、開く前に残存記録を書く。開いた複製を閉じることができたら消す。
        Write-JsonFile $pending @{schema_version=1; run_id=$RunId; process_id=$PID; mode='Render';
            label=[string]$request.label; input=$inputPath; output=$outputPath; work_dir=$WorkDir;
            started_at=[DateTime]::UtcNow.ToString('o'); status='running'; holder=$holderData}
        $recorded = $true
        if ($null -eq $script:app) { $started = New-App }
        try {
            $presentation = $script:app.Presentations.Open($inputPath, -1, 0, 0)
        } catch {
            if (-not (Test-Disconnected $_.Exception)) { throw }
            # 合間に PowerPoint が終了していたので、作り直して 1 回だけやり直す。
            Reset-App
            $started = New-App
            $presentation = $script:app.Presentations.Open($inputPath, -1, 0, 0)
        }
        $presentation.SaveAs($outputPath, 32)
        if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) { throw 'PDF が作成されませんでした。' }
        $result = @{id=$id; ok=$true; renderer='powerpoint'; version=[string]$script:app.Version;
                    build=[string]$script:app.Build; windows_user_sid=$sid; run_id=$RunId;
                    powerpoint_started=$started}
    } catch {
        $result = @{id=$id; ok=$false; renderer='powerpoint'; message=$_.Exception.Message;
                    state_dir=$StateDir; work_dir=$WorkDir; run_id=$RunId}
    }
    # 開いた複製だけを閉じる。利用者の PowerPoint は強制終了しない。
    $cleanup = 'ok'
    if ($null -ne $presentation) {
        try {
            $presentation.Close()
            [Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
        } catch {
            $cleanup = 'failed'
        }
        $presentation = $null
    }
    if ($recorded -and $cleanup -eq 'ok') {
        try { Remove-Item -LiteralPath $pending -Force } catch { $cleanup = 'failed' }
    }
    $result.cleanup = $cleanup
    if ($cleanup -ne 'ok') {
        $result.ok = $false
        $result.message = "終了を確認できません。記録: $pending"
    }
    return $result
}

function Test-ProcessAlive($processId) {
    # 残存記録の作業プロセス（PowerShell）が動いているか。番号がなければ動いていないとみなす。
    if ($null -eq $processId) { return $false }
    try { $number = [int]$processId } catch { return $false }
    if ($number -le 0) { return $false }
    $process = Get-Process -Id $number -ErrorAction SilentlyContinue
    return ($null -ne $process -and $process.ProcessName -match '^(powershell|pwsh)$')
}

function Test-FileLocked([string]$path) {
    # 排他で開くことができなければ真。ファイルがなければ $null。
    if (-not $path) { return $null }
    try {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    } catch {
        return $null
    }
    try {
        $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
        $stream.Close()
        return $false
    } catch {
        return $true
    }
}

function Get-WorkRoot {
    return [IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) 'slidekit')).TrimEnd('\')
}

function Test-WorkFolder([string]$path) {
    # Windows の一時フォルダーの slidekit の直下にある作業フォルダー ppt-* か。
    if (-not $path) { return $false }
    try {
        $full = [IO.Path]::GetFullPath($path).TrimEnd('\')
        $parent = [IO.Path]::GetDirectoryName($full)
        return ($parent -ieq (Get-WorkRoot)) -and ((Split-Path -Leaf $full) -like 'ppt-*')
    } catch {
        return $false
    }
}

function Remove-OldWorkFolders([string]$skip) {
    # 24 時間より古く、どの記録も指していない作業フォルダー ppt-* を消す。消した数を返す。
    $root = Get-WorkRoot
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return 0 }
    $keep = @()
    foreach ($path in @($WorkDir, $skip, (Read-JsonFile $active).work_dir, (Read-JsonFile $pending).work_dir)) {
        if ($path) {
            try { $keep += [IO.Path]::GetFullPath([string]$path).TrimEnd('\') } catch { }
        }
    }
    $limit = [DateTime]::UtcNow.AddHours(-24)
    $count = 0
    foreach ($folder in @(Get-ChildItem -LiteralPath $root -Directory -Filter 'ppt-*' -ErrorAction SilentlyContinue)) {
        if ($keep -contains $folder.FullName.TrimEnd('\')) { continue }
        if ($folder.LastWriteTimeUtc -gt $limit) { continue }
        try {
            Remove-Item -LiteralPath $folder.FullName -Recurse -Force
            $count += 1
        } catch { }
    }
    return $count
}

function Invoke-Clear {
    # 排他を取った状態で呼ぶ。解除の条件を満たさなければ理由を返して断る。
    if (-not (Test-Path -LiteralPath $pending -PathType Leaf)) {
        $removed = Remove-OldWorkFolders ''
        return @{ok=$true; status='nothing'; reason='残存記録はありません。'; archived=$null;
                 removed_work_dir=$false; removed_old_work_dirs=$removed}
    }
    $record = Read-JsonFile $pending
    $processId = $null
    $inputPath = ''
    $workPath = ''
    if ($null -ne $record) {
        $processId = $record.process_id
        $inputPath = [string]$record.input
        $workPath = [string]$record.work_dir
    }
    $refused = @{ok=$true; status='refused'; pending=$record; archived=$null; removed_work_dir=$false;
                 removed_old_work_dirs=0}
    if (Test-ProcessAlive $processId) {
        $refused.reason = "残存記録の処理（PID $processId）が動いています。終わってから再実行してください。"
        return $refused
    }
    $inputLocked = Test-FileLocked $inputPath
    if ($inputLocked -eq $true) {
        $refused.reason = 'PowerPoint が検査用の複製を開いています。PowerPoint で閉じてから再実行してください。'
        return $refused
    }
    # PowerPoint は読み取り専用で開いた複製のファイルを開いたままにしないので、排他で開くことができても、複製を開いていないとは言えない。
    # PowerPoint が起動していれば、複製の有無によらず --force を求める。
    $running = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
    if ($running -and (-not $Force)) {
        $refused.reason = 'PowerPoint が起動しています。PowerPoint が検査用の複製を開いているかは確かめられません。PowerPoint の終了を待つか、複製を開いていないことを確かめてから --force を付けて再実行してください。'
        return $refused
    }
    $archive = Join-Path $StateDir 'cleared'
    New-Item -ItemType Directory -Force -Path $archive | Out-Null
    $target = Join-Path $archive ([DateTime]::UtcNow.ToString("yyyyMMdd'T'HHmmssfff'Z'") + '-pending.json')
    Move-Item -LiteralPath $pending -Destination $target
    $removedWork = $false
    if ((Test-WorkFolder $workPath) -and (Test-Path -LiteralPath $workPath -PathType Container)) {
        try {
            Remove-Item -LiteralPath $workPath -Recurse -Force
            $removedWork = $true
        } catch { }
    }
    $removed = Remove-OldWorkFolders $workPath
    return @{ok=$true; status='cleared'; reason='残存記録を解除しました。'; pending=$record; archived=$target;
             removed_work_dir=$removedWork; removed_old_work_dirs=$removed}
}

function Remove-OwnWorkFolder {
    # 自分の残存記録が残っていなければ、自分の作業フォルダーを消す（切り離された後も片付ける）。
    if (-not (Test-WorkFolder $WorkDir)) { return }
    if (Test-Path -LiteralPath $pending) {
        $record = Read-JsonFile $pending
        if ($null -eq $record -or $record.run_id -eq $RunId) { return }
    }
    try { Remove-Item -LiteralPath $WorkDir -Recurse -Force } catch { }
}

if ($Mode -eq 'Status') {
    $exitCode = 0
    try {
        $busy = -not (Lock-Mutex $false)
        # 取ることができたら直ちに放す。
        Unlock-Mutex
        $current = $null
        if ($busy) { $current = Read-JsonFile $active }
        $pendingExists = Test-Path -LiteralPath $pending -PathType Leaf
        $record = Read-JsonFile $pending
        $alive = $null
        $inputLocked = $null
        if ($pendingExists) {
            $processId = $null
            $inputPath = ''
            if ($null -ne $record) {
                $processId = $record.process_id
                $inputPath = [string]$record.input
            }
            $alive = Test-ProcessAlive $processId
            $inputLocked = Test-FileLocked $inputPath
        }
        $running = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
        # abandoned：排他を持っていた処理が放さずに終わっていた（取って放したので、後の処理には残らない）。
        Send-Result @{ok=$true; busy=$busy; holder=$current; pending=$record; pending_exists=$pendingExists;
                      pending_process_alive=$alive; staged_input_locked=$inputLocked; powerpoint_running=$running;
                      abandoned=$abandoned; state_dir=$StateDir}
    } catch {
        Send-Result @{ok=$false; message=$_.Exception.Message; state_dir=$StateDir}
        $exitCode = 2
    } finally {
        Unlock-Mutex
        $mutex.Dispose()
    }
    exit $exitCode
}

if ($Mode -eq 'Clear') {
    $exitCode = 0
    try {
        if (Lock-Mutex $false) {
            Send-Result (Invoke-Clear)
        } else {
            $current = Read-JsonFile $active
            Send-Result @{ok=$true; status='refused'; reason=('変換の実行中のため解除できません' + (Format-Holder $current) + '。');
                          holder=$current; archived=$null; removed_work_dir=$false; removed_old_work_dirs=0}
        }
    } catch {
        Send-Result @{ok=$false; message=$_.Exception.Message; state_dir=$StateDir}
        $exitCode = 2
    } finally {
        Unlock-Mutex
        $mutex.Dispose()
    }
    exit $exitCode
}

# Session
if ($Holder) {
    try {
        $holderData = ConvertFrom-Json -InputObject ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Holder)))
    } catch {
        $holderData = $null
    }
}
$activeWritten = $false
$closeId = $null
$quit = $false
$cleanupWarning = $null
$exitCode = 0
try {
    # 排他を待つ。最初と以後 5 秒ごとに、使用中の処理の情報を付けて waiting を出す。
    $waitStarted = [DateTime]::UtcNow
    $lastNotice = $null
    $acquired = Lock-Mutex $false
    while ((-not $acquired) -and (([DateTime]::UtcNow - $waitStarted).TotalSeconds -lt $WaitSeconds)) {
        if (($null -eq $lastNotice) -or (([DateTime]::UtcNow - $lastNotice).TotalSeconds -ge 5)) {
            $waited = [Math]::Round(([DateTime]::UtcNow - $waitStarted).TotalSeconds, 1)
            Send-Event @{event='waiting'; waited_seconds=$waited; holder=(Read-JsonFile $active)}
            $lastNotice = [DateTime]::UtcNow
        }
        $acquired = Lock-Mutex $true
    }
    $waited = [Math]::Round(([DateTime]::UtcNow - $waitStarted).TotalSeconds, 1)
    if (-not $acquired) {
        $current = Read-JsonFile $active
        $message = 'SLIDEKIT_BUSY: 同じ Windows 利用者による変換が実行中です' + (Format-Holder $current) + '。'
        if ($WaitSeconds -gt 0) { $message += "$waited 秒待ちました。" }
        Send-Result @{id=0; ok=$false; renderer='powerpoint'; message=$message; holder=$current; waited_seconds=$waited}
        $exitCode = 2
    } else {
        New-Item -ItemType Directory -Force -Path $StateDir | Out-Null
        # 使用中の記録。待つ側が、使用中の処理の情報として示す。排他を放す直前に消す。
        Write-JsonFile $active @{schema_version=1; pid=$PID; run_id=$RunId; started_at=[DateTime]::UtcNow.ToString('o');
                                 work_dir=$WorkDir; holder=$holderData}
        $activeWritten = $true
        if (Test-Path -LiteralPath $pending) {
            Send-Result @{id=0; ok=$false; renderer='powerpoint';
                          message="SLIDEKIT_PENDING: 前回の処理の終了を確認してください。記録: $pending";
                          pending=(Read-JsonFile $pending)}
            $exitCode = 2
        } else {
            Send-Event @{event='ready'; pid=$PID; sid=$sid; state_dir=$StateDir; waited_seconds=$waited;
                         abandoned=$abandoned}
            while ($true) {
                $line = $null
                try { $line = [Console]::In.ReadLine() } catch { $line = $null }
                if ($null -eq $line) { break }
                if ($line.Trim().Length -eq 0) { continue }
                $request = $null
                try { $request = ConvertFrom-Json -InputObject $line } catch { $request = $null }
                if ($null -eq $request) {
                    Send-Result @{id=$null; ok=$false; message='要求を解釈できません。'}
                    continue
                }
                if ($request.op -eq 'close') {
                    $closeId = $request.id
                    break
                }
                if ($request.op -eq 'doctor') {
                    Send-Result (Invoke-Doctor $request.id)
                } elseif ($request.op -eq 'render') {
                    Send-Result (Invoke-Render $request)
                } else {
                    Send-Result @{id=$request.id; ok=$false; message="未対応の要求です: $($request.op)"}
                }
            }
        }
    }
} catch {
    Send-Result @{id=0; ok=$false; renderer='powerpoint'; message=$_.Exception.Message; state_dir=$StateDir;
                  work_dir=$WorkDir; run_id=$RunId}
    $exitCode = 2
} finally {
    # 後片付け：自分で起動した PowerPoint だけを、開いた資料がなければ終了する。
    if ($null -ne $app) {
        try {
            if ((-not $existing) -and $app.Presentations.Count -eq 0) {
                $app.Quit()
                $quit = $true
            }
        } catch {
            $cleanupWarning = $_.Exception.Message
        }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
        $app = $null
    }
    if ($activeWritten) {
        try { Remove-Item -LiteralPath $active -Force } catch { }
    }
    Unlock-Mutex
    $mutex.Dispose()
    if ($null -ne $closeId) {
        Send-Result @{id=$closeId; ok=$true; quit=$quit; cleanup_warning=$cleanupWarning}
    }
    Remove-OwnWorkFolder
}
exit $exitCode
