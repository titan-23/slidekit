param(
    [ValidateSet('Session', 'Status', 'Clear')][string]$Mode,
    [string]$StateDir,
    [string]$WorkDir,
    [string]$RunId,
    [double]$WaitSeconds = 0,
    [string]$Holder = '',
    [switch]$Force,
    [switch]$StopOwned
)

# slidekit の PowerPoint の作業スクリプト。Windows 利用者ごとの排他で、変換を一度に一つにする。
# Session：排他を取り、標準入力のリクエスト（1 行に 1 つの JSON）に応答する。標準入力の終了は close と同じ扱い。
# Status：排他・残存記録・使用中の処理・PowerPoint の処理を調べる。Clear：安全を確かめてから残存記録を解除する。
# 利用者の PowerPoint は強制終了しない。開いた複製だけを閉じ、自分で起動して資料が残っていない場合だけ終了する。
# 強制終了するのは、Clear の -StopOwned で、slidekit が起動したと記録（PID と開始時刻）で確かめることができ、
# ウィンドウのない PowerPoint だけである（Stop-OwnedPowerPoint）。
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
# この作業プロセスが使う PowerPoint（New-App が決める）。自分で起動したら started と PID・開始時刻を持つ。
$powerpoint = $null
# 使用中の記録（active.json）の内容。PowerPoint を起動したら powerpoint を加えて書き直す。
$activeRecord = $null
# 前の処理が消さずに残した使用中の記録にある、slidekit が起動した PowerPoint（Get-OwnedRecords の形）。
$previous = @()
# 変換の途中か。複製を閉じることができず残存記録を残したか。
$rendering = $false
$keepPending = $false

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

function Get-StartTime($process) {
    # 処理の開始時刻。start_time は UTC の FILETIME の整数（照合に使う）、started_at は表示用の文字列。取得できなければ $null。
    try {
        $time = $process.StartTime
        if ($null -eq $time) { return $null }
        return @{start_time=$time.ToFileTimeUtc(); started_at=$time.ToUniversalTime().ToString('o')}
    } catch {
        return $null
    }
}

function Get-PowerPointInfo($process) {
    # PowerPoint の処理の PID・ウィンドウの題名・表示中のウィンドウの有無・開始時刻。
    # 取得できない値は $null になる（PowerShell はプロパティの取得の例外を $null にする）。ウィンドウが分からなければ、あるとみなす。
    $title = ''
    $window = $true
    try { $title = [string]$process.MainWindowTitle } catch { }
    try {
        $handle = $process.MainWindowHandle
        $window = ($null -eq $handle) -or ($handle -ne [IntPtr]::Zero)
    } catch { }
    $info = @{pid=$process.Id; window_title=$title; has_window=($window -or $title.Length -gt 0);
              started_at=$null; start_time=$null}
    $start = Get-StartTime $process
    if ($null -ne $start) {
        $info.started_at = $start.started_at
        $info.start_time = $start.start_time
    }
    return $info
}

function Get-OwnedRecords {
    # slidekit が起動した PowerPoint の記録（残存記録と使用中の記録の powerpoint）。PID と開始時刻があるものだけを返す。
    # 使用中の記録は、排他を持つ処理がなければ、前の処理が消さずに残したものである。
    $list = @()
    foreach ($item in @(@{source='pending'; path=$pending}, @{source='active'; path=$active})) {
        $record = Read-JsonFile $item.path
        if ($null -eq $record -or $null -eq $record.powerpoint) { continue }
        $owned = $record.powerpoint
        if ($owned.started -ne $true -or $null -eq $owned.pid -or $null -eq $owned.start_time) { continue }
        try {
            # 表示用の開始時刻は整数から作る（PowerShell 7 の ConvertFrom-Json は日時の文字列を日時の値に変えるため）。
            $time = [long]$owned.start_time
            $list += @{source=$item.source; run_id=[string]$record.run_id; pid=[int]$owned.pid; start_time=$time;
                       started_at=[DateTime]::FromFileTimeUtc($time).ToString('o')}
        } catch { }
    }
    return $list
}

function Find-OwnedRecord($info, $records) {
    # PID と開始時刻の両方が一致する記録を返す。開始時刻を取得できない処理は一致としない。
    if ($null -eq $info.start_time) { return $null }
    foreach ($record in $records) {
        if ($record.pid -eq $info.pid -and $record.start_time -eq [long]$info.start_time) { return $record }
    }
    return $null
}

function Test-RecordedPowerPointGone($record) {
    # 残存記録の変換に使った PowerPoint が slidekit の起動したものと記録されていて、もう動いていなければ真。
    # このときは、ほかの PowerPoint が検査用の複製を開いていることはない（複製はその PowerPoint で開いた）。
    if ($null -eq $record -or $null -eq $record.powerpoint) { return $false }
    $owned = $record.powerpoint
    if ($owned.started -ne $true -or $null -eq $owned.pid -or $null -eq $owned.start_time) { return $false }
    try {
        $target = @{pid=[int]$owned.pid; start_time=[long]$owned.start_time}
    } catch {
        return $false
    }
    foreach ($process in @(Get-Process POWERPNT -ErrorAction SilentlyContinue)) {
        if ($process.Id -ne $target.pid) { continue }
        $start = Get-StartTime $process
        # 開始時刻を取得できなければ、同じ処理かもしれないので、動いているとみなす。
        if ($null -eq $start -or [long]$start.start_time -eq $target.start_time) { return $false }
    }
    return $true
}

function New-App {
    # 既に動いている PowerPoint があれば、それに接続する（終了させない）。自分で起動したかを返す。
    # 前の処理が終了させている途中の PowerPoint があると、COM は新しい PowerPoint を起動するので、PowerPoint があったかでなく、
    # 接続の前後の処理の番号を比べて、新しい処理が増えたときだけ自分で起動したとする。
    $before = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $script:app = New-Object -ComObject PowerPoint.Application
    $after = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
    $added = @($after | Where-Object { $before -notcontains $_.Id })
    $script:existing = $added.Count -eq 0
    # 起動した PowerPoint の PID と開始時刻を記録し、後の renderer status と renderer clear --stop-owned が照合できるようにする。
    # 増えた処理が 1 つでなければ、どれを起動したかを決めることができないので、PID を記録しない。
    $script:powerpoint = @{started=(-not $script:existing); pid=$null; started_at=$null; start_time=$null}
    if ($added.Count -eq 1) {
        $start = Get-StartTime $added[0]
        if ($null -ne $start) {
            $script:powerpoint.pid = $added[0].Id
            $script:powerpoint.started_at = $start.started_at
            $script:powerpoint.start_time = $start.start_time
        }
    } elseif ($added.Count -eq 0 -and $before.Count -eq 1 -and @($script:previous).Count -gt 0) {
        # 動いている PowerPoint が 1 つだけで、前の処理が起動して終了させずに残したもの（記録と PID・開始時刻が一致し、
        # ウィンドウがない）なら、接続先はそれである。自分で起動したものとして引き継ぎ、最後に資料がなければ終了する。
        $only = @($after | Where-Object { $_.Id -eq $before[0] })
        if ($only.Count -eq 1) {
            $info = Get-PowerPointInfo $only[0]
            if ((-not $info.has_window) -and $null -ne (Find-OwnedRecord $info $script:previous)) {
                $script:existing = $false
                $script:powerpoint = @{started=$true; pid=$info.pid; started_at=$info.started_at;
                                       start_time=$info.start_time; inherited=$true}
            }
        }
    }
    if ($null -ne $script:activeRecord) {
        $script:activeRecord.powerpoint = $script:powerpoint
        try { Write-JsonFile $active $script:activeRecord } catch { }
    }
    # 戻り値（生成記録の powerpoint_started）は、PowerPoint の処理が増えたかで決める（引き継いだものは含めない）。
    return ($added.Count -gt 0)
}

function Reset-App {
    if ($null -ne $script:app) {
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($script:app) | Out-Null } catch { }
    }
    $script:app = $null
    $script:powerpoint = $null
}

function Write-Pending($fields) {
    # 残存記録を書く。powerpoint は変換に使う PowerPoint（起動の前なら $null）。
    # Python 側が書き換えた状態（timeout・interrupted）は保つ。
    $record = @{schema_version=1; run_id=$RunId; process_id=$PID; mode='Render'; label=$fields.label;
                input=$fields.input; output=$fields.output; work_dir=$WorkDir; started_at=$fields.started_at;
                status='running'; holder=$holderData; powerpoint=$script:powerpoint}
    $current = Read-JsonFile $pending
    if ($null -ne $current -and $current.run_id -eq $RunId -and $current.status) { $record.status = [string]$current.status }
    Write-JsonFile $pending $record
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
        # PowerPoint を起動したら、その PID と開始時刻を加えて書き直してから開く。
        $fields = @{label=[string]$request.label; input=$inputPath; output=$outputPath;
                    started_at=[DateTime]::UtcNow.ToString('o')}
        $script:rendering = $true
        Write-Pending $fields
        $recorded = $true
        if ($null -eq $script:app) {
            $started = New-App
            Write-Pending $fields
        }
        try {
            $presentation = $script:app.Presentations.Open($inputPath, -1, 0, 0)
        } catch {
            if (-not (Test-Disconnected $_.Exception)) { throw }
            # 合間に PowerPoint が終了していたので、作り直して 1 回だけやり直す。
            Reset-App
            $started = New-App
            Write-Pending $fields
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
    if ($recorded -and $cleanup -ne 'ok') { $script:keepPending = $true }
    $script:rendering = $false
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

function Get-OwnedProcesses($records) {
    # 記録と PID・開始時刻の両方が一致する PowerPoint の処理。
    # 照合と終了が同じ処理を指すよう、照合の前にハンドルを開いておく（開くことができなければ opened が偽）。
    # PowerShell はプロパティの取得の例外を $null にするので、例外でなく値で判断する。
    $list = @()
    if (@($records).Count -eq 0) { return $list }
    foreach ($process in @(Get-Process POWERPNT -ErrorAction SilentlyContinue)) {
        $handle = $null
        try { $handle = $process.Handle } catch { }
        $opened = ($null -ne $handle) -and ($handle -ne [IntPtr]::Zero)
        $info = Get-PowerPointInfo $process
        $match = Find-OwnedRecord $info $records
        if ($null -ne $match) { $list += @{process=$process; info=$info; record=$match; opened=$opened} }
    }
    return $list
}

function Stop-OwnedPowerPoint($records) {
    # slidekit が起動したと確かめることができる PowerPoint だけを強制終了する（Clear の -StopOwned）。
    # 条件は、記録と PID・開始時刻の両方が一致し、表示中のウィンドウも題名もないこと。利用者の PowerPoint は終了しない。
    # 戻り値は、終了した処理（stopped）と、断ったときの理由（reason）。
    $stopped = @()
    $owned = @(Get-OwnedProcesses $records)
    foreach ($item in $owned) {
        if ($item.info.has_window) {
            $title = ''
            if ($item.info.window_title) { $title = "（題名「$($item.info.window_title)」）" }
            return @{stopped=$stopped; reason=("PowerPoint（PID $($item.info.pid)）は slidekit が起動したものですが、ウィンドウ$($title)があるので終了しません。" +
                '利用者が資料を開いていることがあります。資料を保存して PowerPoint を終了してから再実行してください。')}
        }
        if (-not $item.opened) {
            return @{stopped=$stopped; reason="PowerPoint（PID $($item.info.pid)）は slidekit が起動したものですが、終了する権限がありません。"}
        }
    }
    foreach ($item in $owned) {
        try { $item.process.Kill() } catch { }
        $exited = $false
        try { $exited = $item.process.WaitForExit(10000) } catch { }
        if (-not $exited) {
            return @{stopped=$stopped; reason=("PowerPoint（PID $($item.info.pid)）の終了を確認できません。" +
                'slidekit renderer status で確かめてから再実行してください。')}
        }
        $stopped += @{pid=$item.info.pid; started_at=$item.info.started_at; record=$item.record.source}
    }
    return @{stopped=$stopped; reason=$null}
}

function Get-UnprovenReason($record) {
    # -StopOwned でも終了する PowerPoint を決めることができない理由。
    $owned = $null
    if ($null -ne $record) { $owned = $record.powerpoint }
    if ($null -eq $owned) {
        $text = '残存記録に、変換に使った PowerPoint の PID と開始時刻がありません（古い形式の記録か、PowerPoint を起動する前に止まった記録）。'
    } elseif ($owned.started -ne $true) {
        $text = '残存記録の変換は、既に起動していた PowerPoint を使いました。slidekit が起動したものではないので終了しません。'
    } elseif ($null -eq $owned.pid) {
        $text = '変換の開始時に PowerPoint の処理が複数増えたため、slidekit が起動した PowerPoint の PID を記録していません。'
    } else {
        $text = "残存記録の PowerPoint（PID $($owned.pid)）の開始時刻を記録と照合することができません。"
    }
    return $text + 'slidekit renderer status で PowerPoint の一覧（PID・ウィンドウの題名・開始時刻）を確かめてください。'
}

function Invoke-Clear {
    # 排他を取った状態で呼ぶ。解除の条件を満たさなければ理由を返して断る。
    # -StopOwned では、作業プロセスが動いていないことを確かめた後、slidekit が起動した PowerPoint だけを終了してから解除する。
    $records = @(Get-OwnedRecords)
    $stopped = @()
    if (-not (Test-Path -LiteralPath $pending -PathType Leaf)) {
        $reason = '残存記録はありません。'
        if ($StopOwned) {
            # 前の処理が終了させずに残した PowerPoint（使用中の記録と一致するもの）を終了する。
            $stop = Stop-OwnedPowerPoint $records
            $stopped = @($stop.stopped)
            if ($null -ne $stop.reason) {
                return @{ok=$true; status='refused'; reason=$stop.reason; archived=$null; removed_work_dir=$false;
                         removed_old_work_dirs=0; stopped=$stopped}
            }
        }
        if ($stopped.Count -gt 0) {
            $reason += 'slidekit が起動した PowerPoint（PID ' + (($stopped | ForEach-Object { $_.pid }) -join '、') + '）を終了しました。'
        }
        $removed = Remove-OldWorkFolders ''
        return @{ok=$true; status='nothing'; reason=$reason; archived=$null;
                 removed_work_dir=$false; removed_old_work_dirs=$removed; stopped=$stopped}
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
                 removed_old_work_dirs=0; stopped=$stopped}
    if (Test-ProcessAlive $processId) {
        $refused.reason = "残存記録の処理（PID $processId）が動いています。終わってから再実行してください。"
        return $refused
    }
    if ($StopOwned) {
        $stop = Stop-OwnedPowerPoint $records
        $stopped = @($stop.stopped)
        $refused.stopped = $stopped
        if ($null -ne $stop.reason) {
            $refused.reason = $stop.reason
            return $refused
        }
    }
    $inputLocked = Test-FileLocked $inputPath
    if ($inputLocked -eq $true) {
        $refused.reason = 'PowerPoint が検査用の複製を開いています。PowerPoint で閉じてから再実行してください。'
        return $refused
    }
    # PowerPoint は読み取り専用で開いた複製のファイルを開いたままにしないので、排他で開くことができても、複製を開いていないとは言えない。
    # PowerPoint が起動していれば、複製の有無によらず --force を求める。ただし、残存記録の変換に使った PowerPoint が
    # slidekit の起動したものと記録されていて、もう動いていなければ、ほかの PowerPoint は複製を開いていないので求めない。
    $running = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
    if ($running -and (-not $Force) -and (-not (Test-RecordedPowerPointGone $record))) {
        $reason = 'PowerPoint が起動しています。PowerPoint が検査用の複製を開いているかは確かめられません。'
        if ($StopOwned) {
            $reason += Get-UnprovenReason $record
        } else {
            $owned = @(Get-OwnedProcesses $records | Where-Object { -not $_.info.has_window })
            if ($owned.Count -gt 0) {
                $reason += 'slidekit が起動した PowerPoint（PID ' + (($owned | ForEach-Object { $_.info.pid }) -join '、') +
                    '、ウィンドウなし）は、--stop-owned を付けると終了してから解除します。'
            }
        }
        $refused.reason = $reason + 'PowerPoint の終了を待つか、複製を開いていないことを確かめてから --force を付けて再実行してください。'
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
    $reason = '残存記録を解除しました。'
    if ($stopped.Count -gt 0) {
        $reason = 'slidekit が起動した PowerPoint（PID ' + (($stopped | ForEach-Object { $_.pid }) -join '、') + '）を終了し、' + $reason
    }
    return @{ok=$true; status='cleared'; reason=$reason; pending=$record; archived=$target;
             removed_work_dir=$removedWork; removed_old_work_dirs=$removed; stopped=$stopped}
}

function Remove-OwnPending {
    # この実行（RunId）の残存記録だけを消す。ほかの実行の記録には触れない。
    if (-not (Test-Path -LiteralPath $pending -PathType Leaf)) { return }
    $record = Read-JsonFile $pending
    if ($null -ne $record -and $record.run_id -eq $RunId) {
        try { Remove-Item -LiteralPath $pending -Force } catch { }
    }
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
        # PowerPoint の処理ごとに、slidekit が起動したと記録した PowerPoint と PID・開始時刻が一致するかを示す。
        $records = @(Get-OwnedRecords)
        $infos = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { Get-PowerPointInfo $_ })
        $processes = @()
        foreach ($info in $infos) {
            $match = Find-OwnedRecord $info $records
            $source = $null
            if ($null -ne $match) { $source = $match.source }
            $processes += @{pid=$info.pid; window_title=$info.window_title; has_window=$info.has_window;
                            started_at=$info.started_at; owned=($null -ne $match); record=$source}
        }
        $recorded = @()
        foreach ($item in $records) {
            $live = @($infos | Where-Object {
                $_.pid -eq $item.pid -and $null -ne $_.start_time -and [long]$_.start_time -eq $item.start_time }).Count -gt 0
            $recorded += @{source=$item.source; run_id=$item.run_id; pid=$item.pid; started_at=$item.started_at; running=$live}
        }
        $running = $infos.Count -gt 0
        # abandoned：排他を持っていた処理が放さずに終わっていた（取って放したので、後の処理には残らない）。
        Send-Result @{ok=$true; busy=$busy; holder=$current; pending=$record; pending_exists=$pendingExists;
                      pending_process_alive=$alive; staged_input_locked=$inputLocked; powerpoint_running=$running;
                      powerpoint_processes=$processes; powerpoint_records=$recorded;
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
        if (Test-Path -LiteralPath $pending) {
            # 使用中の記録は書かない（前の処理が残した記録の PowerPoint の情報を、renderer clear --stop-owned のために残す）。
            Send-Result @{id=0; ok=$false; renderer='powerpoint';
                          message="SLIDEKIT_PENDING: 前回の処理の終了を確認してください。記録: $pending";
                          pending=(Read-JsonFile $pending)}
            $exitCode = 2
        } else {
            # 前の処理が消さずに残した使用中の記録に、slidekit が起動した PowerPoint があれば New-App が引き継ぐ。
            $previous = @(Get-OwnedRecords | Where-Object { $_.source -eq 'active' })
            # 使用中の記録。待つ側が、使用中の処理の情報として示す。排他を放す直前に消す。
            $activeRecord = @{schema_version=1; pid=$PID; run_id=$RunId; started_at=[DateTime]::UtcNow.ToString('o');
                              work_dir=$WorkDir; holder=$holderData}
            Write-JsonFile $active $activeRecord
            $activeWritten = $true
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
    # 切り離された後に Python 側が書いた、この実行の残存記録を消す。変換は終わっているので、後続の変換を止める理由がない。
    # 変換の途中で止まった場合と、開いた複製を閉じることができなかった場合は残す。
    if ($locked -and (-not $rendering) -and (-not $keepPending)) { Remove-OwnPending }
    # 自分で起動した PowerPoint を終了できなかったときは、使用中の記録を消さずに残し、PID と開始時刻を
    # renderer status と renderer clear --stop-owned が照合できるようにする（排他を放した後の使用中の記録は、残した記録とみなす）。
    $keepActive = (-not $existing) -and (-not $quit) -and ($null -ne $powerpoint) -and ($null -ne $powerpoint.pid)
    if ($activeWritten -and (-not $keepActive)) {
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
