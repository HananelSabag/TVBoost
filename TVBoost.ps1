#requires -Version 5.1
<#
  TVBoost — אבחון, ניקוי וייעול לטלוויזיית Xiaomi Mi TV 4S (MiTV-MSSP3, Android TV 9)

  רץ על המחשב ומדבר עם הטלוויזיה דרך ADB ברשת הביתית.
  לא מתקין כלום על הטלוויזיה, לא דורש root ולא מוחק כלום:
  אפליקציות "זבל" רק מושבתות, וכל שינוי נרשם ב-data\undo-*.json וחוזר בכפתור "שחזור הכל".

  TVBoost.ps1             פותח את החלון
  TVBoost.ps1 -SelfTest   בדיקות פענוח + טעינת החלון בלי להציג אותו
#>
param([switch]$SelfTest, [string]$RenderPng)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$DataDir = Join-Path $Root 'data'
$ReportDir = Join-Path $Root 'reports'
foreach ($d in $DataDir, $ReportDir) { if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null } }

if (-not ('TvRow' -as [type])) {
  Add-Type -TypeDefinition @'
public class TvRow {
  public bool Sel { get; set; }
  public string Pkg { get; set; }
  public string Name { get; set; }
  public string Info { get; set; }
  public string State { get; set; }
  public string MB { get; set; }
  public string Tag { get; set; }
}
'@
}

function Find-Adb {
  $cands = @()
  $cmd = Get-Command adb.exe -ErrorAction SilentlyContinue
  if ($cmd) { $cands += $cmd.Source }
  $cands += (Join-Path $Root 'platform-tools\adb.exe')
  $cands += (Join-Path $env:LOCALAPPDATA 'Android\Sdk\platform-tools\adb.exe')
  if ($env:ANDROID_HOME) { $cands += (Join-Path $env:ANDROID_HOME 'platform-tools\adb.exe') }
  foreach ($c in $cands) { if ($c -and (Test-Path $c)) { return $c } }
  $null
}
$AdbPath = Find-Adb

# ---------------------------------------------------------------- knowledge

# אפליקציות שאפשר לכבות בבטחה. D = מסומן כברירת מחדל.
$Catalog = @(
  @{ P = 'com.google.android.tvrecommendations'; N = 'המלצות במסך הבית'; I = 'שורות "מומלץ עבורך" ותוכן מוצע. רץ ברקע כל הזמן ומוריד תמונות.'; D = $true }
  @{ P = 'com.mitv.tvhome.atv'; N = 'PatchWall'; I = 'מסך הבית של שיאומי עם ההמלצות והפרסומות. מסך הבית של גוגל (או Projectivy) נשאר.'; D = $true }
  @{ P = 'com.mitv.tvhome.mitvplus'; N = 'Mi TV+'; I = 'ערוצים ופרסומות של שיאומי.'; D = $true }
  @{ P = 'com.mitv.tvhome.michannel'; N = 'ערוצי ההמלצות של שיאומי'; I = 'שורות "מומלץ" של שיאומי.'; D = $true }
  @{ P = 'com.google.android.backdrop'; N = 'שומר מסך Ambient'; I = 'שומר המסך עם תמונות מגוגל. מוריד תמונות ברקע.'; D = $true }
  @{ P = 'com.google.android.play.games'; N = 'Google Play Games'; I = 'משחקים — לא בשימוש בטלוויזיה.'; D = $true }
  @{ P = 'com.google.android.videos'; N = 'Google Play סרטים'; I = 'שירות ישן של גוגל. תופס זיכרון.'; D = $true }
  @{ P = 'com.google.android.music'; N = 'Google Play Music'; I = 'שירות שגוגל סגרה ב-2020.'; D = $true }
  @{ P = 'com.google.android.feedback'; N = 'משוב לגוגל'; I = 'שליחת דיווחי תקלות לגוגל.'; D = $true }
  @{ P = 'com.google.android.marvin.talkback'; N = 'TalkBack'; I = 'הקראת מסך ללקויי ראייה.'; D = $true }
  @{ P = 'com.google.android.tvtutorials'; N = 'הדרכות Android TV'; I = 'סרטוני הדרכה של גוגל.'; D = $true }
  @{ P = 'com.android.printspooler'; N = 'שירות הדפסה'; I = 'אין מדפסת בטלוויזיה.'; D = $true }
  @{ P = 'com.google.android.youtube.tvmusic'; N = 'YouTube Music'; I = 'כבה רק אם לא שומעים מוזיקה באפליקציה הזו.'; D = $false }
  @{ P = 'com.google.android.youtube.tvkids'; N = 'YouTube Kids'; I = 'כבה רק אם אין ילדים שמשתמשים בה.'; D = $false }
  @{ P = 'com.amazon.amazonvideo.livingroom'; N = 'Prime Video'; I = 'כבה רק אם אין לכם מנוי.'; D = $false }
  @{ P = 'com.google.android.katniss'; N = 'Google Assistant / חיפוש קולי'; I = 'הכי כבד בזיכרון, אבל אחרי כיבוי כפתור המיקרופון בשלט לא יעבוד.'; D = $false }
  @{ P = 'com.google.android.speech.pumpkin'; N = 'זיהוי דיבור'; I = 'חלק מהחיפוש הקולי.'; D = $false }
  @{ P = 'com.google.android.apps.mediashell'; N = 'Chromecast מובנה'; I = 'רץ תמיד ברקע. כבה רק אם לא משדרים מהטלפון לטלוויזיה.'; D = $false }
  @{ P = 'com.google.android.tv.remote.service'; N = 'שלט מהטלפון'; I = 'שליטה בטלוויזיה מאפליקציית Google TV בטלפון.'; D = $false }
)

# אסור לגעת: מערכת, חנות, שירותי גוגל, מקלדת, עדכונים, רכיבי הטלוויזיה עצמה.
$ProtectRx = '^(android|com\.android\.(systemui|settings|tv\.settings|shell|packageinstaller|bluetooth|phone|webview|externalstorage|keychain|vending|tv|providers\..+|inputmethod.*|server.*|networkstack.*|wifi.*|location.*)|com\.google\.android\.(gms.*|gsf.*|webview|tvlauncher|leanback\.ime|packageinstaller|youtube\.tv|tvsetup|setupwraith|partnersetup|configupdater|ext\..+|tv\.frameworkpackagestubs)|com\.netflix\.ninja)$'
# מנהל המשימות לא סוגר את אלה (ליבת המערכת).
$CoreRx = '^(android|system|com\.android\.(systemui|phone|bluetooth|shell|providers\..+|server.*|networkstack.*|inputmethod.*)|com\.google\.android\.(gms.*|gsf.*|leanback\.ime))$'
$ProtectWordRx = '(update|upgrade|ota|ime|input|keyboard|setup|hdmi|tvinput|tvservice|mstar|mediatek|mtk|remote|bluetooth|drm|widevine|playready|keystone|factory|launcher|provider|systemui|settings|permission)'

# אפליקציות מערכת שבטוח לסגור (force-stop) בניקוי — הן יעלו שוב כשצריך.
$SafeStop = @('com.google.android.katniss', 'com.google.android.backdrop', 'com.google.android.play.games', 'com.google.android.videos',
  'com.google.android.youtube.tvmusic', 'com.google.android.youtube.tvkids', 'com.netflix.ninja', 'com.amazon.amazonvideo.livingroom',
  'com.android.vending', 'com.google.android.tvrecommendations', 'com.google.android.youtube.tv')

# לעולם לא לסגור בניקוי.
$AlwaysKeep = @('com.google.android.tvlauncher', 'com.spocky.projengmenu', 'me.efesser.flauncher', 'com.google.android.leanback.ime', 'com.android.systemui', 'com.google.android.gms')

$FriendlyNames = @{
  'com.google.android.youtube.tv' = 'YouTube'; 'com.google.android.tvlauncher' = 'מסך הבית המקורי'; 'com.google.android.katniss' = 'Google Assistant'
  'com.google.android.gms' = 'שירותי Google'; 'com.android.vending' = 'Google Play חנות'; 'com.google.android.apps.mediashell' = 'Chromecast'
  'com.netflix.ninja' = 'Netflix'; 'com.spocky.projengmenu' = 'Projectivy (מסך בית)'; 'me.efesser.flauncher' = 'FLauncher (מסך בית)'
  'com.google.android.tvrecommendations' = 'המלצות'; 'com.google.android.backdrop' = 'שומר מסך'; 'system' = 'מערכת Android'
  'com.android.systemui' = 'ממשק מערכת'; 'com.google.android.youtube.tvmusic' = 'YouTube Music'; 'com.google.android.leanback.ime' = 'מקלדת'
}

$Launchers = @(
  @{ P = 'com.spocky.projengmenu'; N = 'Projectivy Launcher'; I = 'מומלץ. קל, בלי פרסומות והמלצות, בוחרים בדיוק אילו אפליקציות מופיעות (YouTube, Netflix, yes+...).' }
  @{ P = 'me.efesser.flauncher'; N = 'FLauncher'; I = 'קוד פתוח, מינימלי לגמרי. פחות אפשרויות.' }
  @{ P = 'com.google.android.tvlauncher'; N = 'מסך הבית המקורי'; I = 'של גוגל, עם שורות ההמלצות.' }
)

# ---------------------------------------------------------------- helpers (also run in the background runspace)

function L { param([string]$T, [string]$K = 'info') [pscustomobject]@{ T = $T; K = $K } }

function Invoke-Adb {
  param([string[]]$A, [int]$TimeoutSec = 60)
  if (-not $AdbPath) { throw 'adb.exe לא נמצא במחשב' }
  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = $AdbPath
  $psi.Arguments = (($A | ForEach-Object { if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ } }) -join ' ')
  $psi.UseShellExecute = $false
  $psi.RedirectStandardOutput = $true
  $psi.RedirectStandardError = $true
  $psi.CreateNoWindow = $true
  $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
  $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
  $p = [System.Diagnostics.Process]::Start($psi)
  $o = $p.StandardOutput.ReadToEndAsync()
  $e = $p.StandardError.ReadToEndAsync()
  if (-not $p.WaitForExit($TimeoutSec * 1000)) { try { $p.Kill() } catch { }; return [pscustomobject]@{ Code = -1; Out = ''; Err = 'timeout' } }
  $out = ''; $err = ''
  if ($o.Wait(5000)) { $out = $o.Result }
  if ($e.Wait(2000)) { $err = $e.Result }
  [pscustomobject]@{ Code = $p.ExitCode; Out = ($out -replace "`r", ''); Err = ($err -replace "`r", '') }
}

# פקודת shell בטלוויזיה המחוברת ($Serial מגיע מה-scope של המשימה)
function Tv {
  param([string]$Cmd, [int]$TimeoutSec = 60)
  $r = Invoke-Adb @('-s', $Serial, 'shell', $Cmd) $TimeoutSec
  if ($r.Code -ne 0 -and -not $r.Out -and $r.Err -match "device|offline|closed|not found|unauthorized|timeout") {
    throw "החיבור לטלוויזיה נפל ($($r.Err.Trim())). לחץ התחבר שוב."
  }
  $r.Out
}

function Get-Sections {
  param([string]$Raw)
  $h = @{}; $cur = '_'
  $h[$cur] = New-Object System.Collections.Generic.List[string]
  foreach ($line in ($Raw -split "`n")) {
    if ($line -match '^@@(\w+)\s*$') { $cur = $Matches[1]; $h[$cur] = New-Object System.Collections.Generic.List[string]; continue }
    $h[$cur].Add($line)
  }
  $out = @{}
  foreach ($k in $h.Keys) { $out[$k] = (($h[$k]) -join "`n").Trim() }
  $out
}

function ConvertFrom-ProcMeminfo {
  param([string]$Text)
  $m = @{}
  foreach ($line in ($Text -split "`n")) { if ($line -match '^(\w+):\s+(\d+)\s*kB') { $m[$Matches[1]] = [long]$Matches[2] } }
  $m
}

function ConvertFrom-DumpsysMeminfo {
  param([string]$Text)
  $procs = New-Object System.Collections.Generic.List[object]
  $status = ''; $in = $false
  foreach ($line in ($Text -split "`n")) {
    if ($line -match '^Total PSS by process') { $in = $true; continue }
    if ($in) {
      if ($line -match '^\s*([\d,]+)K:\s+(\S+)\s+\(pid\s+(\d+)') {
        $procs.Add([pscustomobject]@{ Name = $Matches[2]; ProcId = [int]$Matches[3]; MB = [math]::Round([double]($Matches[1] -replace ',', '') / 1024, 0) })
      } elseif ($procs.Count -gt 0 -and $line -match '^\S') { $in = $false }
    }
    if ($line -match '^\s*Total RAM:.*status (\w+)') { $status = $Matches[1] }
  }
  [pscustomobject]@{ Procs = $procs; Status = $status }
}

function ConvertFrom-Ping {
  param([string]$Text)
  $r = [pscustomobject]@{ Sent = 0; Recv = 0; Loss = 100; Avg = $null; Max = $null; Mdev = $null; Ip = '' }
  if ($Text -match '(\d+) packets transmitted, (\d+) received') {
    $r.Sent = [int]$Matches[1]; $r.Recv = [int]$Matches[2]
    if ($r.Sent) { $r.Loss = [math]::Round(100 * ($r.Sent - $r.Recv) / $r.Sent) }
  }
  if ($Text -match '= ([\d.]+)/([\d.]+)/([\d.]+)/([\d.]+) ms') { $r.Avg = [double]$Matches[2]; $r.Max = [double]$Matches[3]; $r.Mdev = [double]$Matches[4] }
  if ($Text -match 'PING \S+ \(([\d.]+)\)') { $r.Ip = $Matches[1] }
  $r
}

function ConvertFrom-WifiInfo {
  param([string]$Text)
  $w = [pscustomobject]@{ Ssid = ''; Rssi = $null; Link = $null; Freq = $null; State = '' }
  if ($Text -match '(?<!B)SSID: "?([^",]*)"?,') { $w.Ssid = $Matches[1] }
  if ($Text -match 'Supplicant state: (\w+)') { $w.State = $Matches[1] }
  if ($Text -match 'RSSI: (-?\d+)') { $w.Rssi = [int]$Matches[1] }
  if ($Text -match 'Link speed: (\d+)') { $w.Link = [int]$Matches[1] }
  if ($Text -match 'Frequency: (\d+)') { $w.Freq = [int]$Matches[1] }
  $w
}

function ConvertFrom-Df {
  param([string]$Text)
  foreach ($line in ($Text -split "`n")) {
    if ($line -match '^\S+\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)%') {
      return [pscustomobject]@{ SizeKB = [long]$Matches[1]; UsedKB = [long]$Matches[2]; AvailKB = [long]$Matches[3]; UsePct = [int]$Matches[4] }
    }
  }
  $null
}

function Get-CpuPct {
  param([string]$A, [string]$B)
  # שמות משתנים ב-PowerShell לא תלויי רישיות: $a היה דורס את הפרמטר $A
  $s1 = @(($A -split '\s+') | Select-Object -Skip 1 | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [long]$_ })
  $s2 = @(($B -split '\s+') | Select-Object -Skip 1 | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [long]$_ })
  if ($s1.Count -lt 5 -or $s2.Count -lt 5) { return $null }
  $dt = ($s2 | Measure-Object -Sum).Sum - ($s1 | Measure-Object -Sum).Sum
  if ($dt -le 0) { return $null }
  $idle = ($s2[3] + $s2[4]) - ($s1[3] + $s1[4])
  [math]::Round(100 * (1 - $idle / $dt))
}

function Get-PkgList {
  param([string]$Flags = '')
  $t = Tv "pm list packages $Flags"
  @(($t -split "`n") | Where-Object { $_ -like 'package:*' } | ForEach-Object { $_.Substring(8).Trim() })
}

function New-PkgSet {
  param([string[]]$Items)
  $s = New-Object 'System.Collections.Generic.HashSet[string]'
  foreach ($i in $Items) { if ($i) { [void]$s.Add($i) } }
  , $s
}

function Resolve-ProcPkg {
  param([string]$Proc, $PkgSet)
  $n = ($Proc -split ':')[0]
  while ($n) {
    if ($PkgSet.Contains($n)) { return $n }
    $i = $n.LastIndexOf('.')
    if ($i -lt 0) { break }
    $n = $n.Substring(0, $i)
  }
  $null
}

function Get-RunningPkgs {
  param($PkgSet)
  $names = (Tv 'ps -A -o NAME') -split "`n" | Select-Object -Skip 1
  $set = New-Object 'System.Collections.Generic.HashSet[string]'
  foreach ($n in $names) { $p = Resolve-ProcPkg $n.Trim() $PkgSet; if ($p) { [void]$set.Add($p) } }
  , $set
}

function Get-TvHome {
  $t = Tv 'cmd package resolve-activity --brief -a android.intent.action.MAIN -c android.intent.category.HOME'
  $last = ($t -split "`n" | Where-Object { $_ -match '^\s*[\w.]+/[\w.$]+\s*$' } | Select-Object -Last 1)
  if ($last) { return $last.Trim() }
  ''
}

# מה אסור לסגור עכשיו: האפליקציה שעל המסך, מסך הבית, המקלדת
function Get-TvKeepSet {
  $keep = New-PkgSet $AlwaysKeep
  $f = Tv "dumpsys activity activities | grep -m 1 -E 'mResumedActivity|mFocusedActivity'"
  if ($f -match 'u0 ([\w.]+)/') { [void]$keep.Add($Matches[1]) }
  $h = Get-TvHome
  if ($h) { [void]$keep.Add(($h -split '/')[0]) }
  $ime = (Tv 'settings get secure default_input_method').Trim()
  if ($ime -match '^([\w.]+)/') { [void]$keep.Add($Matches[1]) }
  , $keep
}

function Get-FriendlyName {
  param([string]$Pkg)
  if ($FriendlyNames.ContainsKey($Pkg)) { return $FriendlyNames[$Pkg] }
  foreach ($c in $Catalog) { if ($c.P -eq $Pkg) { return $c.N } }
  $Pkg
}

function Find-TvOnLan {
  $prefixes = @(Get-NetIPConfiguration -ErrorAction SilentlyContinue |
    Where-Object { $_.IPv4DefaultGateway -and $_.NetAdapter.Status -eq 'Up' } |
    ForEach-Object { $_.IPv4Address | ForEach-Object { ($_.IPAddress -split '\.')[0..2] -join '.' } } | Select-Object -Unique)
  $found = @{}
  foreach ($pre in $prefixes) {
    $probes = New-Object System.Collections.Generic.List[object]
    foreach ($i in 1..254) {
      foreach ($port in 5555, 8008, 8009) {
        $c = New-Object System.Net.Sockets.TcpClient
        $probes.Add(@($c, $c.ConnectAsync("$pre.$i", $port), "$pre.$i", $port))
      }
    }
    Start-Sleep -Milliseconds 1800
    foreach ($pr in $probes) {
      if ($pr[1].Status -eq 'RanToCompletion') {
        if (-not $found.ContainsKey($pr[2])) { $found[$pr[2]] = @{ Adb = $false; Cast = $false } }
        if ($pr[3] -eq 5555) { $found[$pr[2]].Adb = $true } else { $found[$pr[2]].Cast = $true }
      }
      $pr[0].Dispose()
    }
  }
  foreach ($ip in $found.Keys) {
    $name = ''
    if ($found[$ip].Cast) { try { $name = (Invoke-RestMethod -Uri "http://$ip`:8008/setup/eureka_info?params=name" -TimeoutSec 2).name } catch { } }
    [pscustomobject]@{ Ip = $ip; Adb = $found[$ip].Adb; Name = $name }
  }
}

function Connect-Tv {
  param([string]$Ip)
  $target = $Ip
  if ($target -notmatch ':\d+$') { $target = "$Ip`:5555" }
  $r = Invoke-Adb @('connect', $target) 20
  $msg = ($r.Out + ' ' + $r.Err).Trim()
  $state = ''
  for ($i = 0; $i -lt 30; $i++) {
    $s = Invoke-Adb @('-s', $target, 'get-state') 8
    $state = ($s.Out + ' ' + $s.Err).Trim()
    if ($state -eq 'device') { break }
    if ($state -notmatch 'unauthorized|authorizing' -and $i -lt 2) { $msg = (Invoke-Adb @('connect', $target) 15).Out.Trim() }
    if ($state -notmatch 'unauthorized|authorizing' -and $i -ge 3) { break }
    Start-Sleep -Seconds 2
  }
  if ($state -ne 'device') { return [pscustomobject]@{ Ok = $false; Target = $target; Msg = $msg; State = $state } }
  $Serial = $target
  $p = @((Tv 'getprop ro.product.model; getprop ro.build.version.release; getprop ro.build.display.id; getprop ro.serialno') -split "`n" | ForEach-Object { $_.Trim() })
  while ($p.Count -lt 4) { $p += '' }
  [pscustomobject]@{ Ok = $true; Target = $target; Model = $p[0]; Android = $p[1]; Build = $p[2]; SerialNo = $p[3] }
}

function Add-Issue { param($List, [int]$Sev, [string]$Text) $List.Add([pscustomobject]@{ S = $Sev; T = $Text }) }

function Get-TvSystemReport {
  $lines = New-Object System.Collections.Generic.List[object]
  $issues = New-Object System.Collections.Generic.List[object]
  $raw = Tv ('echo @@props; for k in ro.product.model ro.build.version.release ro.build.display.id ro.build.version.security_patch ro.product.device; do getprop $k; done; ' +
    'echo @@uptime; cat /proc/uptime; echo @@mem; cat /proc/meminfo; echo @@df; df -k /data; ' +
    'echo @@cpu1; head -n 1 /proc/stat; sleep 1; echo @@cpu2; head -n 1 /proc/stat; ' +
    'echo @@temp; cat /sys/class/thermal/thermal_zone*/temp 2>/dev/null; echo @@top; top -b -n 2 -d 1 -m 8 2>/dev/null') 60
  $s = Get-Sections $raw
  $p = @(($s['props'] -split "`n") | ForEach-Object { $_.Trim() }); while ($p.Count -lt 5) { $p += '' }

  $lines.Add((L 'מערכת' 'head'))
  $lines.Add((L "דגם: $($p[0]) ($($p[4])) · Android $($p[1]) · גרסה $($p[2]) · עדכון אבטחה $($p[3])"))

  $up = 0.0
  if ($s['uptime'] -match '^([\d.]+)') { $up = [double]$Matches[1] }
  $days = [math]::Floor($up / 86400); $hrs = [math]::Floor(($up % 86400) / 3600)
  $upText = "דולקת ברצף מאז ההפעלה האמיתית האחרונה: $days ימים ו-$hrs שעות"
  if ($up -gt 2 * 86400) {
    $lines.Add((L "$upText — כיבוי בשלט רק מכניס להמתנה, הזיכרון לא מתנקה." 'warn'))
    Add-Issue $issues 1 'לעשות הפעלה מחדש (כפתור "הפעלה מחדש"). מומלץ פעם ביום-יומיים: לחיצה ארוכה על כפתור ההפעלה בשלט ← הפעלה מחדש.'
  } else { $lines.Add((L $upText 'ok')) }

  $m = ConvertFrom-ProcMeminfo $s['mem']
  $totalMB = [math]::Round($m['MemTotal'] / 1024); $availMB = [math]::Round($m['MemAvailable'] / 1024)
  $swapUsedMB = [math]::Round(($m['SwapTotal'] - $m['SwapFree']) / 1024)
  $memText = "זיכרון RAM: $totalMB MB סה""כ · פנוי בפועל $availMB MB · דחוס (zram) בשימוש $swapUsedMB MB"
  if ($availMB -lt 250) { $lines.Add((L $memText 'bad')); Add-Issue $issues 2 'הזיכרון כמעט מלא — לעשות "ניקוי זיכרון" ולכבות אפליקציות זבל ("הסרת זבל").' }
  elseif ($availMB -lt 450) { $lines.Add((L $memText 'warn')); Add-Issue $issues 1 'מעט זיכרון פנוי — "ניקוי זיכרון" ו"הסרת זבל" יעזרו.' }
  else { $lines.Add((L $memText 'ok')) }

  $df = ConvertFrom-Df $s['df']
  if ($df) {
    $freeGB = [math]::Round($df.AvailKB / 1MB, 2); $sizeGB = [math]::Round($df.SizeKB / 1MB, 1)
    $dfText = "אחסון: פנוי $freeGB GB מתוך $sizeGB GB ($($df.UsePct)% בשימוש)"
    if ($freeGB -lt 1) { $lines.Add((L $dfText 'bad')); Add-Issue $issues 2 'האחסון כמעט מלא — זה מאט מאוד. "ניקוי זיכרון" ינקה גם קבצי מטמון; כדאי למחוק אפליקציות שלא בשימוש.' }
    elseif ($freeGB -lt 2) { $lines.Add((L $dfText 'warn')); Add-Issue $issues 1 'אחסון נמוך — כדאי לפנות אפליקציות שלא בשימוש.' }
    else { $lines.Add((L $dfText 'ok')) }
  }

  $cpu = Get-CpuPct $s['cpu1'] $s['cpu2']
  $temps = @(($s['temp'] -split "`n") | Where-Object { $_ -match '^\d+$' } | ForEach-Object { $v = [double]$_; if ($v -gt 1000) { $v / 1000 } else { $v } })
  $tMax = $null; if ($temps.Count) { $tMax = [math]::Round(($temps | Measure-Object -Maximum).Maximum) }
  if ($null -ne $cpu) {
    $cpuText = "מעבד: עומס $cpu%"
    if ($null -ne $tMax) { $cpuText += " · טמפרטורה $tMax°C" }
    if ($cpu -gt 70) { $lines.Add((L "$cpuText — משהו טוחן את המעבד ברקע" 'warn')); Add-Issue $issues 1 'המעבד עמוס גם במנוחה — לפתוח "מנהל משימות" ולסגור את מה שזולל.' }
    else { $lines.Add((L $cpuText 'ok')) }
  }

  $mi = ConvertFrom-DumpsysMeminfo (Tv 'dumpsys meminfo' 90)
  if ($mi.Status) {
    $st = @{ normal = 'תקין'; moderate = 'בינוני'; low = 'נמוך'; critical = 'קריטי' }[$mi.Status]
    $kind = @{ normal = 'ok'; moderate = 'warn'; low = 'bad'; critical = 'bad' }[$mi.Status]
    if (-not $kind) { $kind = 'info'; $st = $mi.Status }
    $lines.Add((L "מצב זיכרון לפי Android: $st" $kind))
  }
  if ($mi.Procs.Count) {
    $lines.Add((L 'מה תופס הכי הרבה זיכרון עכשיו:' 'info'))
    $all = New-PkgSet (Get-PkgList '')
    $agg = @{}
    foreach ($pr in $mi.Procs) { $k = Resolve-ProcPkg $pr.Name $all; if (-not $k) { $k = $pr.Name }; if (-not $agg.ContainsKey($k)) { $agg[$k] = 0 }; $agg[$k] += $pr.MB }
    $agg.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 10 | ForEach-Object {
      $lines.Add((L ("   {0,4} MB  {1}" -f $_.Value, (Get-FriendlyName $_.Key)) 'dim'))
    }
  }

  # שתי דגימות של top — לוקחים את האחרונה (בראשונה האחוזים לא אמינים)
  $topLines = @(($s['top'] -split "`n") | Where-Object { $_ -match '^\s*\d+\s+\S+' } | Select-Object -Last 8)
  if ($topLines.Count) {
    $lines.Add((L 'מה עובד על המעבד:' 'info'))
    foreach ($t in ($topLines | Select-Object -First 6)) {
      $f = @($t.Trim() -split '\s+')
      if ($f.Count -ge 12) {
        $procName = (($f[11..($f.Count - 1)] -join ' ') -split ':')[0]
        $lines.Add((L ("   {0,5}%  {1}" -f $f[8], (Get-FriendlyName $procName)) 'dim'))
      }
    }
  }

  $third = Get-PkgList '-3'
  $dis = Get-PkgList '-d'
  $inst = New-PkgSet (Get-PkgList '')
  $bloatOn = @($Catalog | Where-Object { $_.D -and $inst.Contains($_.P) -and ($dis -notcontains $_.P) })
  $lines.Add((L "אפליקציות שהותקנו: $($third.Count) · רכיבים מושבתים: $($dis.Count)" 'info'))
  if ($bloatOn.Count) {
    $lines.Add((L "יש $($bloatOn.Count) רכיבי זבל פעילים: $(($bloatOn | ForEach-Object { $_.N }) -join ', ')" 'warn'))
    Add-Issue $issues 1 "לכבות $($bloatOn.Count) רכיבי זבל בכפתור ""הסרת זבל"" (הפיך)."
  }
  $homePkg = ((Get-TvHome) -split '/')[0]
  if ($homePkg -notin 'com.spocky.projengmenu', 'me.efesser.flauncher') { Add-Issue $issues 0 'מסך בית נקי (Projectivy) במקום מסך הבית של גוגל — פחות זיכרון, בלי המלצות. כפתור "מסך בית נקי".' }
  [pscustomobject]@{ Lines = $lines; Issues = $issues }
}

function Test-TvLanSpeed {
  $f = Join-Path $env:TEMP 'tvboost-speed.bin'
  if (-not (Test-Path $f) -or (Get-Item $f).Length -ne 16MB) {
    $b = New-Object byte[] (16MB); (New-Object System.Random).NextBytes($b); [IO.File]::WriteAllBytes($f, $b)
  }
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $r = Invoke-Adb @('-s', $Serial, 'push', $f, '/data/local/tmp/tvboost-speed.bin') 120
  $sw.Stop()
  [void](Tv 'rm -f /data/local/tmp/tvboost-speed.bin')
  if ($r.Code -ne 0) { return $null }
  $sec = $sw.Elapsed.TotalSeconds
  if (($r.Out + $r.Err) -match '\((\d+) bytes in ([\d.]+)s\)') { $sec = [double]$Matches[2] }
  if ($sec -le 0) { return $null }
  [math]::Round(16 * 1.048576 * 8 / $sec, 1)
}

function Get-TvNetworkReport {
  $lines = New-Object System.Collections.Generic.List[object]
  $issues = New-Object System.Collections.Generic.List[object]
  $lines.Add((L 'רשת ואינטרנט' 'head'))
  $raw = Tv ('echo @@wifi; dumpsys wifi | grep -m 1 mWifiInfo; echo @@route; ip route get 8.8.8.8; ' +
    'echo @@dns; dumpsys connectivity | grep -m 1 DnsAddresses') 30
  $s = Get-Sections $raw
  $gw = ''; $dev = ''
  if ($s['route'] -match 'via (\S+) dev (\S+)') { $gw = $Matches[1]; $dev = $Matches[2] }
  elseif ($s['route'] -match 'dev (\S+)') { $dev = $Matches[1] }
  $w = ConvertFrom-WifiInfo $s['wifi']

  if ($dev -like 'eth*') { $lines.Add((L 'חיבור: כבל רשת (Ethernet) — מצוין.' 'ok')) }
  elseif ($w.State -eq 'COMPLETED') {
    $band = '?'; $ch = '?'
    if ($w.Freq) { if ($w.Freq -lt 3000) { $band = '2.4GHz'; $ch = [math]::Round(($w.Freq - 2407) / 5) } else { $band = '5GHz'; $ch = [math]::Round(($w.Freq - 5000) / 5) } }
    $lines.Add((L "חיבור: Wi-Fi ""$($w.Ssid)"" · רצועה $band (ערוץ $ch) · עוצמת אות $($w.Rssi) dBm · מהירות קישור $($w.Link) Mbps"))
    if ($null -ne $w.Rssi) {
      if ($w.Rssi -ge -60) { $lines.Add((L 'עוצמת האות טובה.' 'ok')) }
      elseif ($w.Rssi -ge -70) { $lines.Add((L 'עוצמת האות בינונית.' 'warn')) }
      else { $lines.Add((L 'עוצמת האות חלשה — זה מאט סטרימינג.' 'bad')); Add-Issue $issues 2 'האות חלש: לקרב את הנתב/להוסיף מגדיל טווח, או לבדוק שהטלוויזיה לא מאחורי מתכת/קיר.' }
    }
    if ($band -eq '2.4GHz') {
      $lines.Add((L 'הטלוויזיה על 2.4GHz — רצועה איטית ועמוסה (שכנים, מיקרוגל, בלוטות'').' 'warn'))
      Add-Issue $issues 2 'לעבור ל-5GHz: הנתב משדר את שתי הרצועות באותו שם (עם WPA3), אז הטלוויזיה בוחרת לבד ונתקעת על 2.4. בהגדרות הנתב לפצל רשת 5GHz בשם נפרד (WPA2 או WPA2/WPA3 משולב) ולחבר אליה את הטלוויזיה.'
    }
    if ($w.Link -and $w.Link -lt 72) { $lines.Add((L "מהירות הקישור נמוכה ($($w.Link) Mbps)." 'warn')) }
  } else { $lines.Add((L 'לא הצלחתי לקרוא את פרטי ה-Wi-Fi.' 'warn')) }

  if ($s['dns'] -match '\[(.*)\]') { $lines.Add((L "שרתי DNS: $(($Matches[1] -replace '/', '').Trim(' ,'))" 'dim')) }

  if ($gw) {
    $pg = ConvertFrom-Ping (Tv "ping -c 10 -i 0.3 -W 1 $gw" 30)
    $t = "פינג לנתב ($gw): ממוצע $($pg.Avg) ms · קפיצות (jitter) $($pg.Mdev) ms · אובדן $($pg.Loss)%"
    if ($pg.Loss -gt 0 -or $pg.Avg -gt 15 -or $pg.Mdev -gt 10) { $lines.Add((L "$t — הקישור האלחוטי לא יציב" 'bad')); Add-Issue $issues 2 'הקשר בין הטלוויזיה לנתב לא יציב (אובדן/קפיצות). 5GHz בשם נפרד ופחות מכשולים בדרך יעזרו.' }
    else { $lines.Add((L $t 'ok')) }
  }
  $pi = ConvertFrom-Ping (Tv 'ping -c 10 -i 0.3 -W 2 8.8.8.8' 30)
  $t = "פינג לאינטרנט (Google): ממוצע $($pi.Avg) ms · קפיצות $($pi.Mdev) ms · אובדן $($pi.Loss)%"
  if ($pi.Recv -eq 0) { $lines.Add((L 'אין גישה לאינטרנט מהטלוויזיה!' 'bad')); Add-Issue $issues 2 'הטלוויזיה לא מגיעה לאינטרנט — לבדוק חיבור Wi-Fi.' }
  elseif ($pi.Loss -gt 0 -or $pi.Avg -gt 60) { $lines.Add((L $t 'warn')) }
  else { $lines.Add((L $t 'ok')) }

  $dnsOut = Tv 'a=$(date +%s%N); ping -c 1 -W 3 www.youtube.com >/dev/null 2>&1; r=$?; b=$(date +%s%N); echo $r $(( (b-a)/1000000 ))' 15
  if ($dnsOut -match '^(\d+) (\d+)') {
    if ($Matches[1] -eq '0') { $lines.Add((L "DNS + פינג ל-YouTube: $($Matches[2]) ms" 'ok')) }
    else { $lines.Add((L 'YouTube לא ענה לפינג (לפעמים זה תקין — תלוי בשרת).' 'dim')) }
  }

  $speed = Test-TvLanSpeed
  if ($null -ne $speed) {
    $t = "מהירות Wi-Fi בין הנתב לטלוויזיה (נמדד מהמחשב): $speed Mbps"
    if ($speed -lt 25) { $lines.Add((L "$t — לא מספיק ל-4K (צריך 25+)" 'bad')); Add-Issue $issues 2 'מהירות ה-Wi-Fi של הטלוויזיה נמוכה מ-4K — המעבר ל-5GHz הוא התיקון העיקרי.' }
    elseif ($speed -lt 50) { $lines.Add((L "$t — מספיק, אבל בלי מרווח" 'warn')) }
    else { $lines.Add((L $t 'ok')) }
    $lines.Add((L '(הבדיקה עוברת דרך המחשב ב-Wi-Fi, אז המספר האמיתי של הטלוויזיה יכול להיות קצת גבוה יותר.)' 'dim'))
  }
  [pscustomobject]@{ Lines = $lines; Issues = $issues }
}

function Get-TvDiagnosis {
  param([string]$Mode = 'full')
  $out = New-Object System.Collections.Generic.List[object]
  $issues = New-Object System.Collections.Generic.List[object]
  if ($Mode -ne 'net') { $r = Get-TvSystemReport; $out.AddRange([object[]]$r.Lines); foreach ($i in $r.Issues) { $issues.Add($i) } }
  $n = Get-TvNetworkReport; $out.AddRange([object[]]$n.Lines); foreach ($i in $n.Issues) { $issues.Add($i) }
  $out.Add((L 'מה כדאי לעשות' 'head'))
  if ($issues.Count -eq 0) { $out.Add((L 'הכל נראה טוב.' 'ok')) }
  foreach ($i in ($issues | Sort-Object S -Descending)) { $out.Add((L $i.T 'tip')) }
  $out
}

function Invoke-TvClean {
  $out = New-Object System.Collections.Generic.List[object]
  $m1 = ConvertFrom-ProcMeminfo (Tv 'cat /proc/meminfo')
  $all = New-PkgSet (Get-PkgList '')
  $third = New-PkgSet (Get-PkgList '-3')
  $keep = Get-TvKeepSet
  $running = Get-RunningPkgs $all
  $stopped = @()
  foreach ($pkg in $running) {
    if ($keep.Contains($pkg)) { continue }
    if (-not ($third.Contains($pkg) -or $SafeStop -contains $pkg)) { continue }
    [void](Tv "am force-stop $pkg"); $stopped += $pkg
  }
  [void](Tv 'am kill-all')
  $df = ConvertFrom-Df (Tv 'df -k /data')
  if ($df -and $df.AvailKB -lt 1.5 * 1MB) {
    [void](Tv 'pm trim-caches 999G' 180)
    $df2 = ConvertFrom-Df (Tv 'df -k /data')
    if ($df2) { $out.Add((L ("נוקו קבצי מטמון: אחסון פנוי {0:N2} GB ← {1:N2} GB" -f ($df.AvailKB / 1MB), ($df2.AvailKB / 1MB)) 'ok')) }
  }
  Start-Sleep -Seconds 2
  $m2 = ConvertFrom-ProcMeminfo (Tv 'cat /proc/meminfo')
  if ($stopped.Count) { $out.Add((L "נסגרו $($stopped.Count) אפליקציות ברקע: $(($stopped | ForEach-Object { Get-FriendlyName $_ }) -join ', ')" 'ok')) }
  else { $out.Add((L 'לא היו אפליקציות מיותרות פתוחות ברקע.' 'ok')) }
  $a1 = [math]::Round($m1['MemAvailable'] / 1024); $a2 = [math]::Round($m2['MemAvailable'] / 1024)
  $out.Add((L "זיכרון פנוי: $a1 MB ← $a2 MB ($(if ($a2 -ge $a1) { '+' })$($a2 - $a1) MB)" 'ok'))
  $out.Add((L 'האפליקציה שעל המסך, מסך הבית והמקלדת לא נסגרו.' 'dim'))
  $out
}

function Get-TvTaskRows {
  $mi = ConvertFrom-DumpsysMeminfo (Tv 'dumpsys meminfo' 90)
  $all = New-PkgSet (Get-PkgList '')
  $third = New-PkgSet (Get-PkgList '-3')
  $keep = Get-TvKeepSet
  $agg = @{}
  foreach ($pr in $mi.Procs) { $k = Resolve-ProcPkg $pr.Name $all; if (-not $k) { continue }; if (-not $agg.ContainsKey($k)) { $agg[$k] = 0 }; $agg[$k] += $pr.MB }
  $agg.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 40 | ForEach-Object {
    $pkg = $_.Key
    $kind = 'מערכת'
    if ($third.Contains($pkg)) { $kind = 'אפליקציה' }
    if ($pkg -match $CoreRx) { $kind = 'ליבת מערכת' }
    if ($keep.Contains($pkg)) { $kind = 'פתוח עכשיו / מסך בית' }
    [pscustomobject]@{ Sel = $false; Pkg = $pkg; Name = (Get-FriendlyName $pkg); Info = $pkg; State = $kind; MB = "$($_.Value)"; Tag = '' }
  }
}

function Stop-TvApps {
  param([string[]]$Pkgs)
  $done = @()
  foreach ($p in $Pkgs) {
    if ($p -match $CoreRx) { L "לא נסגר (ליבת מערכת): $p" 'warn'; continue }
    [void](Tv "am force-stop $p"); $done += $p
  }
  if ($done.Count) { L "נסגרו: $(($done | ForEach-Object { Get-FriendlyName $_ }) -join ', ')" 'ok' }
}

function Get-TvBloatRows {
  $u = Get-PkgList '-u'
  $inst = New-PkgSet (Get-PkgList '')
  $dis = New-PkgSet (Get-PkgList '-d')
  $sys = Get-PkgList '-s'
  $running = Get-RunningPkgs (New-PkgSet $u)
  $catSet = New-PkgSet ($Catalog | ForEach-Object { $_.P })
  foreach ($c in $Catalog) {
    if ($u -notcontains $c.P) { continue }
    $state = 'פעיל'; $active = $true
    if (-not $inst.Contains($c.P)) { $state = 'הוסר'; $active = $false }
    elseif ($dis.Contains($c.P)) { $state = 'מושבת'; $active = $false }
    elseif ($running.Contains($c.P)) { $state = 'פעיל · רץ עכשיו' }
    [pscustomobject]@{ Sel = ($c.D -and $active); Pkg = $c.P; Name = $c.N; Info = $c.I; State = $state; MB = ''; Tag = 'catalog' }
  }
  foreach ($p in $sys) {
    if ($catSet.Contains($p)) { continue }
    if ($p -notmatch '^com\.(xiaomi|mitv|miui|mi)\.' -and $p -notmatch '(analytics|statistic|tracker|feedback|report)') { continue }
    if ($p -match $ProtectRx -or $p -match $ProtectWordRx) { continue }
    $state = 'פעיל'
    if ($dis.Contains($p)) { $state = 'מושבת' } elseif ($running.Contains($p)) { $state = 'פעיל · רץ עכשיו' }
    [pscustomobject]@{ Sel = $false; Pkg = $p; Name = $p; Info = 'רכיב של שיאומי — לא מסומן. לכבות רק אם יודעים מה הוא.'; State = $state; MB = ''; Tag = 'vendor' }
  }
}

function Invoke-TvDisable {
  param([string[]]$Pkgs)
  foreach ($p in $Pkgs) {
    if ($p -match $ProtectRx) { [pscustomobject]@{ Pkg = $p; Ok = $false; Method = ''; Msg = 'מוגן — לא נוגעים' }; continue }
    $o = Tv "pm disable-user --user 0 $p"
    if ($o -match 'new state: disabled') { [pscustomobject]@{ Pkg = $p; Ok = $true; Method = 'disable'; Msg = '' }; continue }
    $o2 = Tv "pm uninstall -k --user 0 $p"
    if ($o2 -match 'Success') { [pscustomobject]@{ Pkg = $p; Ok = $true; Method = 'uninstall'; Msg = '' }; continue }
    [pscustomobject]@{ Pkg = $p; Ok = $false; Method = ''; Msg = (($o + ' ' + $o2).Trim()) }
  }
}

function Invoke-TvEnable {
  param([object[]]$Entries)
  foreach ($e in $Entries) {
    $ok = $false
    if ($e.Method -ne 'uninstall') { $o = Tv "pm enable $($e.Pkg)"; $ok = $o -match 'new state: enabled' }
    if (-not $ok) { $o = Tv "cmd package install-existing $($e.Pkg)"; $ok = $o -match 'installed for user|Success' ; if ($ok) { [void](Tv "pm enable $($e.Pkg)") } }
    [pscustomobject]@{ Pkg = $e.Pkg; Ok = $ok }
  }
}

$TweakValues = [ordered]@{
  'global window_animation_scale'     = '0.5'
  'global transition_animation_scale' = '0.5'
  'global animator_duration_scale'    = '0.5'
  'global wifi_scan_always_enabled'   = '0'
  'global ble_scan_always_enabled'    = '0'
}

function Invoke-TvTweaks {
  $old = @{}
  foreach ($k in $TweakValues.Keys) {
    $ns, $name = $k -split ' '
    $old[$k] = (Tv "settings get $ns $name").Trim()
    [void](Tv "settings put $ns $name $($TweakValues[$k])")
  }
  [pscustomobject]@{ Old = $old }
}

function Invoke-TvRestore {
  param([object[]]$Packages, [hashtable]$Settings, [string]$HomeComp)
  $out = New-Object System.Collections.Generic.List[object]
  if ($Packages.Count) {
    $res = @(Invoke-TvEnable $Packages)
    $ok = @($res | Where-Object { $_.Ok }).Count
    $out.Add((L "הוחזרו $ok מתוך $($res.Count) אפליקציות." $(if ($ok -eq $res.Count) { 'ok' } else { 'warn' })))
    $out.Add([pscustomobject]@{ T = ''; K = 'restored'; Pkgs = @($res | Where-Object { $_.Ok } | ForEach-Object { $_.Pkg }) })
  }
  if ($Settings -and $Settings.Count) {
    foreach ($k in $Settings.Keys) {
      $parts = @($k -split ' ')
      $old = $Settings[$k]
      $none = (-not $old -or $old -eq 'null')
      if ($parts[0] -eq 'device_config') {
        # "device_config <namespace> <key>" — e.g. the background-app limit set by Shket's stage 1
        if ($none) { [void](Tv "device_config delete $($parts[1]) $($parts[2])") } else { [void](Tv "device_config put $($parts[1]) $($parts[2]) $old") }
      } elseif ($none) { [void](Tv "settings delete $($parts[0]) $($parts[1])") }
      else { [void](Tv "settings put $($parts[0]) $($parts[1]) $old") }
    }
    $out.Add((L 'ההגדרות (אנימציות, סריקות, מגבלת אפליקציות ברקע) הוחזרו למה שהיה.' 'ok'))
  }
  if ($HomeComp) { [void](Tv "cmd package set-home-activity $HomeComp"); $out.Add((L 'מסך הבית הוחזר למקורי.' 'ok')) }
  $out
}

function Get-TvLauncherRows {
  $inst = New-PkgSet (Get-PkgList '')
  $homePkg = ((Get-TvHome) -split '/')[0]
  foreach ($l in $Launchers) {
    $state = 'לא מותקן'
    if ($inst.Contains($l.P)) { $state = 'מותקן' }
    if ($homePkg -eq $l.P) { $state = 'מסך הבית עכשיו' }
    [pscustomobject]@{ Sel = $false; Pkg = $l.P; Name = $l.N; Info = $l.I; State = $state; MB = ''; Tag = 'launcher' }
  }
}

function Open-TvStore {
  param([string]$Pkg)
  [void](Tv "am start -a android.intent.action.VIEW -d 'market://details?id=$Pkg'")
  L "נפתח דף ההתקנה בטלוויזיה — לחץ שם ""התקן"" עם השלט, ואז חזור לכאן ולחץ ""קבע כמסך הבית""." 'tip'
}

function Set-TvHome {
  param([string]$Pkg)
  $prev = Get-TvHome
  $q = Tv 'cmd package query-activities --brief -a android.intent.action.MAIN -c android.intent.category.HOME'
  $comp = $null
  if ($q -match ([regex]::Escape($Pkg) + '/[\w.$]+')) { $comp = $Matches[0] }
  if (-not $comp) {
    $q = Tv 'cmd package query-activities -a android.intent.action.MAIN -c android.intent.category.HOME'
    if ($q -match ('name=(\S+)\s+packageName=' + [regex]::Escape($Pkg))) { $comp = "$Pkg/$($Matches[1])" }
  }
  if (-not $comp) { return [pscustomobject]@{ Ok = $false; Prev = $prev; Msg = 'האפליקציה לא מותקנת (או שאינה מסך בית). קודם "פתח בחנות" והתקן.' } }
  $r = Tv "cmd package set-home-activity $comp"
  [void](Tv 'input keyevent KEYCODE_HOME')
  [pscustomobject]@{ Ok = ($r -match 'Success'); Prev = $prev; Msg = $r.Trim(); Comp = $comp }
}

$BgFunctions = 'L', 'Invoke-Adb', 'Tv', 'Get-Sections', 'ConvertFrom-ProcMeminfo', 'ConvertFrom-DumpsysMeminfo', 'ConvertFrom-Ping',
'ConvertFrom-WifiInfo', 'ConvertFrom-Df', 'Get-CpuPct', 'Get-PkgList', 'New-PkgSet', 'Resolve-ProcPkg', 'Get-RunningPkgs', 'Get-TvHome',
'Get-TvKeepSet', 'Get-FriendlyName', 'Find-TvOnLan', 'Connect-Tv', 'Add-Issue', 'Get-TvSystemReport', 'Test-TvLanSpeed', 'Get-TvNetworkReport',
'Get-TvDiagnosis', 'Invoke-TvClean', 'Get-TvTaskRows', 'Stop-TvApps', 'Get-TvBloatRows', 'Invoke-TvDisable', 'Invoke-TvEnable',
'Invoke-TvTweaks', 'Invoke-TvRestore', 'Get-TvLauncherRows', 'Open-TvStore', 'Set-TvHome'
$BgVariables = 'AdbPath', 'Catalog', 'ProtectRx', 'CoreRx', 'ProtectWordRx', 'SafeStop', 'AlwaysKeep', 'FriendlyNames', 'Launchers', 'TweakValues'

# ---------------------------------------------------------------- self test

function Invoke-SelfTest {
  $script:fail = 0
  function Check($name, $cond) { if ($cond) { Write-Host "ok   $name" } else { Write-Host "FAIL $name" -ForegroundColor Red; $script:fail++ } }
  $mi = ConvertFrom-DumpsysMeminfo @"
Applications Memory Usage (in Kilobytes):
Uptime: 123 Realtime: 123

Total PSS by process:
    185,123K: system (pid 812)
    120,004K: com.google.android.katniss:search (pid 2211 / activities)
     98,000K: com.google.android.youtube.tv (pid 3301 / activities)

Total PSS by OOM adjustment:
    200,000K: Native
 Total RAM: 1,945,060K (status moderate)
"@
  Check 'meminfo procs' ($mi.Procs.Count -eq 3 -and $mi.Procs[1].Name -eq 'com.google.android.katniss:search' -and $mi.Procs[0].MB -eq 181)
  Check 'meminfo status' ($mi.Status -eq 'moderate')
  $set = New-PkgSet @('com.google.android.katniss', 'com.google.android.gms')
  Check 'proc->pkg colon' ((Resolve-ProcPkg 'com.google.android.katniss:search' $set) -eq 'com.google.android.katniss')
  Check 'proc->pkg dotted' ((Resolve-ProcPkg 'com.google.android.gms.persistent' $set) -eq 'com.google.android.gms')
  Check 'proc->pkg none' ($null -eq (Resolve-ProcPkg 'surfaceflinger' $set))
  $w = ConvertFrom-WifiInfo 'mWifiInfo SSID: "home-wifi", BSSID: 02:11:22:33:44:55, MAC: 02:00:00:00:00:00, Supplicant state: COMPLETED, RSSI: -63, Link speed: 72Mbps, Frequency: 2437MHz, Net ID: 0'
  Check 'wifi' ($w.Ssid -eq 'home-wifi' -and $w.Rssi -eq -63 -and $w.Link -eq 72 -and $w.Freq -eq 2437 -and $w.State -eq 'COMPLETED')
  $pg = ConvertFrom-Ping "PING 192.168.1.1 (192.168.1.1) 56(84) bytes of data.`n--- 192.168.1.1 ping statistics ---`n10 packets transmitted, 9 received, 10% packet loss, time 2712ms`nrtt min/avg/max/mdev = 1.234/3.456/12.5/2.1 ms"
  Check 'ping' ($pg.Loss -eq 10 -and $pg.Avg -eq 3.456 -and $pg.Mdev -eq 2.1 -and $pg.Ip -eq '192.168.1.1')
  $df = ConvertFrom-Df "Filesystem      1K-blocks    Used Available Use% Mounted on`n/dev/block/dm-1   4901216 4100000    801216  84% /data"
  Check 'df' ($df.AvailKB -eq 801216 -and $df.UsePct -eq 84)
  Check 'cpu' ((Get-CpuPct 'cpu  100 0 100 700 100 0 0 0 0 0' 'cpu  150 0 150 750 100 0 0 0 0 0') -eq 67)
  $sec = Get-Sections "@@a`nx`ny`n@@b`nz"
  Check 'sections' ($sec['a'] -eq "x`ny" -and $sec['b'] -eq 'z')
  Check 'protect gms' ('com.google.android.gms' -match $ProtectRx)
  Check 'protect vending' ('com.android.vending' -match $ProtectRx)
  Check 'protect youtube' ('com.google.android.youtube.tv' -match $ProtectRx)
  Check 'catalog not protected' (@($Catalog | Where-Object { $_.P -match $ProtectRx }).Count -eq 0)
  Check 'vendor update protected' ('com.xiaomi.mitv.upgrade' -match $ProtectWordRx)
  Check 'adb found' ([bool]$AdbPath)
  $script:fail
}

# ---------------------------------------------------------------- window

[xml]$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="TVBoost — Mi TV" Width="1200" Height="790" MinWidth="960" MinHeight="620"
        WindowStartupLocation="CenterScreen" FlowDirection="RightToLeft"
        Background="#0D1117" Foreground="#E6EAF0" FontFamily="Segoe UI" FontSize="14">
  <Window.Resources>
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Foreground" Value="#E6EAF0"/>
      <Setter Property="Background" Value="#171D26"/>
      <Setter Property="BorderBrush" Value="#263040"/>
      <Setter Property="Padding" Value="14,9"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="B" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="10" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="B" Property="BorderBrush" Value="#3B82F6"/>
                <Setter TargetName="B" Property="Background" Value="#1C2533"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="B" Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Menu" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Margin" Value="0,0,0,7"/>
      <Setter Property="Padding" Value="14,8"/>
    </Style>
    <Style x:Key="Primary" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Background" Value="#2563EB"/>
      <Setter Property="BorderBrush" Value="#2563EB"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="Plain" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
    </Style>
    <Style x:Key="Title" TargetType="TextBlock"><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="FontSize" Value="15"/></Style>
    <Style x:Key="Sub" TargetType="TextBlock"><Setter Property="Foreground" Value="#8A96A8"/><Setter Property="FontSize" Value="12"/></Style>
  </Window.Resources>

  <Grid Margin="20,16,20,12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <DockPanel Grid.Row="0" LastChildFill="False">
      <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">
        <TextBlock Text="TVBoost" FontSize="30" FontWeight="Bold" Foreground="#FFFFFF" FlowDirection="LeftToRight"/>
        <TextBlock Text="אבחון וייעול לטלוויזיית שיאומי" Margin="14,0,0,4" VerticalAlignment="Bottom" Foreground="#8A96A8" FontSize="15"/>
      </StackPanel>
      <Border x:Name="StatusPill" DockPanel.Dock="Right" CornerRadius="14" Padding="14,5" Background="#2A1E1E" VerticalAlignment="Center">
        <TextBlock x:Name="StatusText" Text="לא מחובר" Foreground="#F87171" FontWeight="SemiBold"/>
      </Border>
    </DockPanel>

    <Border Grid.Row="1" Margin="0,14,0,14" Padding="14,10" CornerRadius="12" Background="#131922" BorderBrush="#1F2833" BorderThickness="1">
      <DockPanel LastChildFill="True">
        <TextBlock DockPanel.Dock="Left" Text="כתובת הטלוויזיה:" VerticalAlignment="Center" Margin="0,0,10,0"/>
        <TextBox x:Name="IpBox" DockPanel.Dock="Left" Width="170" FlowDirection="LeftToRight" VerticalContentAlignment="Center" Padding="8,5"
                 Background="#0D1117" Foreground="#E6EAF0" BorderBrush="#263040" CaretBrush="#E6EAF0" FontFamily="Consolas" FontSize="15"/>
        <Button x:Name="BtnScan" DockPanel.Dock="Left" Style="{StaticResource Plain}" Margin="10,0,0,0" Content="חפש טלוויזיה ברשת"/>
        <Button x:Name="BtnConnect" DockPanel.Dock="Left" Style="{StaticResource Primary}" Margin="10,0,0,0" Content="התחבר"/>
        <Button x:Name="BtnHelp" DockPanel.Dock="Left" Style="{StaticResource Plain}" Margin="10,0,0,0" Content="איך מפעילים?"/>
        <TextBlock x:Name="DeviceText" Margin="16,0,0,0" VerticalAlignment="Center" Foreground="#8A96A8" TextTrimming="CharacterEllipsis"/>
      </DockPanel>
    </Border>

    <Grid Grid.Row="2">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="270"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto" Margin="0,0,14,0">
        <StackPanel>
          <Button x:Name="BtnDiag" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="אבחון מלא"/><TextBlock Style="{StaticResource Sub}" Text="זיכרון, מעבד, אחסון, Wi-Fi ואינטרנט"/></StackPanel></Button>
          <Button x:Name="BtnClean" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="ניקוי זיכרון"/><TextBlock Style="{StaticResource Sub}" Text="סוגר את כל מה שרץ ברקע"/></StackPanel></Button>
          <Button x:Name="BtnTasks" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="מנהל משימות"/><TextBlock Style="{StaticResource Sub}" Text="מה תופס זיכרון — ולסגור"/></StackPanel></Button>
          <Button x:Name="BtnBloat" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="הסרת זבל"/><TextBlock Style="{StaticResource Sub}" Text="לכבות המלצות ושירותים מיותרים (הפיך)"/></StackPanel></Button>
          <Button x:Name="BtnLauncher" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="מסך בית נקי"/><TextBlock Style="{StaticResource Sub}" Text="YouTube, Netflix, yes+ — בלי המלצות"/></StackPanel></Button>
          <Button x:Name="BtnTweaks" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="הגדרות מהירות"/><TextBlock Style="{StaticResource Sub}" Text="אנימציות קצרות, פחות סריקות ברקע"/></StackPanel></Button>
          <Button x:Name="BtnNet" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="בדיקת אינטרנט"/><TextBlock Style="{StaticResource Sub}" Text="אות, פינג ומהירות Wi-Fi"/></StackPanel></Button>
          <Button x:Name="BtnReboot" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="הפעלה מחדש"/><TextBlock Style="{StaticResource Sub}" Text="ניקוי מלא של הזיכרון"/></StackPanel></Button>
          <Button x:Name="BtnRestore" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="שחזור הכל"/><TextBlock Style="{StaticResource Sub}" Text="מחזיר כל שינוי שהכלי עשה"/></StackPanel></Button>
          <Button x:Name="BtnSave" Style="{StaticResource Menu}"><StackPanel><TextBlock Style="{StaticResource Title}" Text="שמור דוח"/><TextBlock Style="{StaticResource Sub}" Text="קובץ טקסט עם כל התוצאות"/></StackPanel></Button>
        </StackPanel>
      </ScrollViewer>

      <Border Grid.Column="1" CornerRadius="12" Background="#11161E" BorderBrush="#1F2833" BorderThickness="1">
        <Grid>
          <ScrollViewer x:Name="LogScroll" VerticalScrollBarVisibility="Auto" Padding="18,8,18,14">
            <StackPanel x:Name="LogPanel"/>
          </ScrollViewer>
          <Grid x:Name="ListPanel" Visibility="Collapsed" Margin="16,12,16,14">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <TextBlock x:Name="ListTitle" FontSize="17" FontWeight="SemiBold" Margin="0,0,0,10"/>
            <DataGrid x:Name="Grid" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                      HeadersVisibility="Column" GridLinesVisibility="None" SelectionMode="Single"
                      Background="#11161E" RowBackground="#151B24" AlternatingRowBackground="#121821" Foreground="#E6EAF0"
                      BorderThickness="0" RowHeight="40" FontSize="13">
              <DataGrid.ColumnHeaderStyle>
                <Style TargetType="DataGridColumnHeader">
                  <Setter Property="Background" Value="#0D1117"/>
                  <Setter Property="Foreground" Value="#8A96A8"/>
                  <Setter Property="Padding" Value="8,6"/>
                  <Setter Property="BorderThickness" Value="0"/>
                </Style>
              </DataGrid.ColumnHeaderStyle>
              <DataGrid.CellStyle>
                <Style TargetType="DataGridCell">
                  <Setter Property="BorderThickness" Value="0"/>
                  <Setter Property="Foreground" Value="#E6EAF0"/>
                  <Setter Property="VerticalContentAlignment" Value="Center"/>
                  <Style.Triggers>
                    <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#1E3A8A"/></Trigger>
                  </Style.Triggers>
                </Style>
              </DataGrid.CellStyle>
              <DataGrid.Columns>
                <DataGridTemplateColumn Header="" Width="40">
                  <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate><CheckBox IsChecked="{Binding Sel, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center"/></DataTemplate>
                  </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="שם" Binding="{Binding Name}" Width="210" IsReadOnly="True"/>
                <DataGridTextColumn Header="זיכרון MB" Binding="{Binding MB}" Width="80" IsReadOnly="True"/>
                <DataGridTextColumn Header="מצב" Binding="{Binding State}" Width="120" IsReadOnly="True"/>
                <DataGridTextColumn Header="פרטים" Binding="{Binding Info}" Width="*" IsReadOnly="True">
                  <DataGridTextColumn.ElementStyle><Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/><Setter Property="VerticalAlignment" Value="Center"/></Style></DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
              </DataGrid.Columns>
            </DataGrid>
            <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,12,0,0">
              <Button x:Name="ListBtn1" Style="{StaticResource Primary}" Margin="0,0,10,0"/>
              <Button x:Name="ListBtn2" Style="{StaticResource Plain}" Margin="0,0,10,0"/>
              <Button x:Name="ListBtn3" Style="{StaticResource Plain}" Margin="0,0,10,0"/>
              <Button x:Name="ListClose" Style="{StaticResource Plain}" Content="חזרה ליומן"/>
            </StackPanel>
          </Grid>
        </Grid>
      </Border>
    </Grid>

    <DockPanel Grid.Row="3" Margin="0,10,0,0">
      <TextBlock x:Name="BusyText" DockPanel.Dock="Left" Text="מוכן" Foreground="#8A96A8"/>
      <TextBlock x:Name="FootText" DockPanel.Dock="Right" HorizontalAlignment="Right" Foreground="#4B5563" FontSize="12"/>
    </DockPanel>
  </Grid>
</Window>
'@

$Win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $Xaml))
$ui = @{}
foreach ($n in 'StatusPill', 'StatusText', 'IpBox', 'BtnScan', 'BtnConnect', 'BtnHelp', 'DeviceText', 'BtnDiag', 'BtnClean', 'BtnTasks', 'BtnBloat',
  'BtnLauncher', 'BtnTweaks', 'BtnNet', 'BtnReboot', 'BtnRestore', 'BtnSave', 'LogScroll', 'LogPanel', 'ListPanel', 'ListTitle', 'Grid',
  'ListBtn1', 'ListBtn2', 'ListBtn3', 'ListClose', 'BusyText', 'FootText') { $ui[$n] = $Win.FindName($n) }
$ui.MbCol = $ui.Grid.Columns[2]

$BrushConv = New-Object System.Windows.Media.BrushConverter
function Brush([string]$Hex) { $BrushConv.ConvertFromString($Hex) }
$KindColor = @{ head = '#FFFFFF'; ok = '#4ADE80'; warn = '#FBBF24'; bad = '#F87171'; info = '#D7DEE8'; dim = '#8A96A8'; tip = '#60A5FA' }
$KindPrefix = @{ ok = '✔  '; warn = '⚠  '; bad = '✖  '; tip = '➜  ' }
$script:LogText = New-Object System.Collections.Generic.List[string]

function Write-Log {
  param($Items)
  foreach ($it in @($Items)) {
    if ($null -eq $it) { continue }
    if ($it -is [string]) { $it = L $it }
    if (-not $KindColor.ContainsKey([string]$it.K)) { continue }
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.TextWrapping = 'Wrap'
    $tb.Text = "$($KindPrefix[$it.K])$($it.T)"
    $tb.Foreground = Brush $KindColor[$it.K]
    if ($it.K -eq 'head') {
      $tb.FontSize = 18; $tb.FontWeight = 'SemiBold'; $tb.Margin = '0,14,0,4'
      $script:LogText.Add(''); $script:LogText.Add("== $($it.T) ==")
    } else {
      $tb.Margin = '0,2,0,2'
      if ($it.K -eq 'dim') { $tb.FontSize = 13; $tb.FontFamily = 'Consolas, Segoe UI' }
      $script:LogText.Add("$($KindPrefix[$it.K])$($it.T)")
    }
    [void]$ui.LogPanel.Children.Add($tb)
  }
  $ui.LogScroll.ScrollToEnd()
}

function Set-Connected {
  param([bool]$On, [string]$Text = '')
  if ($On) { $ui.StatusText.Text = 'מחובר'; $ui.StatusText.Foreground = Brush '#4ADE80'; $ui.StatusPill.Background = Brush '#14271C' }
  else { $ui.StatusText.Text = 'לא מחובר'; $ui.StatusText.Foreground = Brush '#F87171'; $ui.StatusPill.Background = Brush '#2A1E1E'; $script:Serial = $null }
  $ui.DeviceText.Text = $Text
}

function Show-AdbHelp {
  Write-Log @(
    (L 'איך מחברים את הטלוויזיה (פעם אחת בלבד)' 'head')
    (L '1. בטלוויזיה: הגדרות (גלגל השיניים) ← העדפות מכשיר ← מידע כללי.')
    (L '2. לרדת לשורה של גרסת ה-build (מתחילה ב-RTT2 או PI, נגמרת ב-release-keys) וללחוץ עליה 7 פעמים ברצף, עד שכתוב "אתה מפתח".')
    (L '3. לחזור אחורה ← "אפשרויות למפתחים" ← להדליק "ניפוי באגים ב-USB" (USB debugging).')
    (L '4. כאן: "חפש טלוויזיה ברשת" ← "התחבר". בטלוויזיה יקפוץ חלון — לסמן "אפשר תמיד מהמחשב הזה" וללחוץ OK.')
    (L 'הטלוויזיה צריכה להיות דלוקה (לא בהמתנה) ובאותה רשת כמו המחשב.' 'dim')
  )
}

# ---------------------------------------------------------------- background tasks

$script:Busy = $false
$script:Current = $null
$script:Serial = $null
$script:DevId = 'tv'
$script:ActionButtons = @($ui.BtnScan, $ui.BtnConnect, $ui.BtnDiag, $ui.BtnClean, $ui.BtnTasks, $ui.BtnBloat, $ui.BtnLauncher, $ui.BtnTweaks,
  $ui.BtnNet, $ui.BtnReboot, $ui.BtnRestore, $ui.ListBtn1, $ui.ListBtn2, $ui.ListBtn3)

function Set-Busy { param([string]$Label) $script:Busy = $true; $ui.BusyText.Text = "עובד: $Label…"; foreach ($b in $script:ActionButtons) { $b.IsEnabled = $false } }
function Set-Idle { $script:Busy = $false; $ui.BusyText.Text = 'מוכן'; foreach ($b in $script:ActionButtons) { $b.IsEnabled = $true } }

$Timer = New-Object System.Windows.Threading.DispatcherTimer
$Timer.Interval = [TimeSpan]::FromMilliseconds(200)
$Timer.Add_Tick({
    $c = $script:Current
    if (-not $c -or -not $c.Handle.IsCompleted) { return }
    $script:Current = $null
    $Timer.Stop()
    $out = $null; $err = $null
    try {
      $out = $c.PS.EndInvoke($c.Handle)
      if ((-not $out -or $out.Count -eq 0) -and $c.PS.Streams.Error.Count) { $err = $c.PS.Streams.Error[0].ToString() }
    } catch {
      $ex = $_.Exception; while ($ex.InnerException) { $ex = $ex.InnerException }; $err = $ex.Message
    }
    $c.PS.Dispose()
    Set-Idle
    if ($err) { Write-Log (L $err 'bad'); if ($err -match 'החיבור לטלוויזיה נפל') { Set-Connected $false } ; return }
    try { & $c.Done $out } catch { Write-Log (L "שגיאה בהצגת התוצאה: $($_.Exception.Message)" 'bad') }
  })

function Start-TvTask {
  param([string]$Label, [scriptblock]$Work, [object[]]$ArgList = @(), [scriptblock]$Done = { param($r) Write-Log $r }, [switch]$NeedTv)
  if ($script:Busy) { return }
  if ($NeedTv -and -not $script:Serial) { Write-Log (L 'קודם צריך להתחבר לטלוויזיה (למעלה: "חפש טלוויזיה ברשת" ← "התחבר").' 'warn'); return }
  Set-Busy $Label
  $ps = [powershell]::Create()
  $ps.RunspacePool = $script:Pool
  [void]$ps.AddScript($Work.ToString())
  foreach ($a in $ArgList) { [void]$ps.AddArgument($a) }
  $script:Current = @{ PS = $ps; Handle = $ps.BeginInvoke(); Done = $Done }
  $Timer.Start()
}

# ---------------------------------------------------------------- undo log

function Get-UndoPath { Join-Path $DataDir ('undo-' + ($script:DevId -replace '[^\w.-]', '_') + '.json') }
function Read-Undo {
  $u = @{ Packages = @(); Settings = @{}; Home = '' }
  $p = Get-UndoPath
  if (Test-Path $p) {
    try {
      $j = Get-Content $p -Raw -Encoding UTF8 | ConvertFrom-Json
      $u.Packages = @($j.Packages | Where-Object { $_ } | ForEach-Object { @{ Pkg = $_.Pkg; Method = $_.Method; When = $_.When } })
      if ($j.Settings) { foreach ($pr in $j.Settings.PSObject.Properties) { $u.Settings[$pr.Name] = $pr.Value } }
      if ($j.Home) { $u.Home = $j.Home }
    } catch { }
  }
  $u
}
function Save-Undo { param($U) ConvertTo-Json -InputObject $U -Depth 5 | Set-Content -Path (Get-UndoPath) -Encoding UTF8 }

function Save-Settings { @{ LastIp = $ui.IpBox.Text.Trim() } | ConvertTo-Json | Set-Content (Join-Path $DataDir 'settings.json') -Encoding UTF8 }

# ---------------------------------------------------------------- list panel

$script:ListMode = ''
function Show-List {
  param([string]$Mode, [string]$Title, $Rows)
  $script:ListMode = $Mode
  $ui.ListTitle.Text = $Title
  $items = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
  foreach ($r in @($Rows)) {
    if ($null -eq $r -or -not $r.Pkg) { continue }
    $row = New-Object TvRow
    $row.Sel = [bool]$r.Sel; $row.Pkg = $r.Pkg; $row.Name = $r.Name; $row.Info = $r.Info; $row.State = $r.State; $row.MB = $r.MB; $row.Tag = $r.Tag
    $items.Add($row)
  }
  $ui.Grid.ItemsSource = $items
  $cfg = @{
    tasks    = @('סגור מסומנים', 'רענן', $null)
    bloat    = @('כבה מסומנים', 'החזר מסומנים', 'רענן')
    launcher = @('פתח בחנות בטלוויזיה', 'קבע כמסך הבית', 'החזר מסך בית מקורי')
  }[$Mode]
  $btns = @($ui.ListBtn1, $ui.ListBtn2, $ui.ListBtn3)
  for ($i = 0; $i -lt 3; $i++) {
    if ($cfg[$i]) { $btns[$i].Content = $cfg[$i]; $btns[$i].Visibility = 'Visible' } else { $btns[$i].Visibility = 'Collapsed' }
  }
  $ui.MbCol.Visibility = $(if ($Mode -eq 'tasks') { 'Visible' } else { 'Collapsed' })
  $ui.LogScroll.Visibility = 'Collapsed'
  $ui.ListPanel.Visibility = 'Visible'
}
function Hide-List { $ui.ListPanel.Visibility = 'Collapsed'; $ui.LogScroll.Visibility = 'Visible'; $ui.LogScroll.ScrollToEnd() }
function Get-Checked { @($ui.Grid.ItemsSource | Where-Object { $_.Sel }) }

function Open-Tasks { Start-TvTask 'טעינת מנהל משימות' { param($Serial) Get-TvTaskRows } @($script:Serial) -NeedTv -Done { param($r) Show-List 'tasks' 'מנהל משימות — מה תופס זיכרון עכשיו (סמן וסגור)' $r } }
function Open-Bloat { Start-TvTask 'סריקת אפליקציות' { param($Serial) Get-TvBloatRows } @($script:Serial) -NeedTv -Done { param($r) Show-List 'bloat' 'הסרת זבל — המסומנים בטוחים לכיבוי. הכל הפיך ("החזר מסומנים" / "שחזור הכל").' $r } }
function Open-Launchers { Start-TvTask 'בדיקת מסך הבית' { param($Serial) Get-TvLauncherRows } @($script:Serial) -NeedTv -Done { param($r) Show-List 'launcher' 'מסך בית נקי — סמן אחד: "פתח בחנות" להתקנה, ואז "קבע כמסך הבית"' $r } }

$ui.ListClose.Add_Click({ Hide-List })

$ui.ListBtn1.Add_Click({
    $sel = Get-Checked
    if (-not $sel.Count) { [void][System.Windows.MessageBox]::Show('לא סומן כלום.', 'TVBoost'); return }
    $pkgs = @($sel | ForEach-Object { $_.Pkg })
    switch ($script:ListMode) {
      'tasks' {
        Start-TvTask 'סגירת אפליקציות' { param($Serial, $P) Stop-TvApps $P } @($script:Serial, $pkgs) -Done { param($r) Write-Log $r; Open-Tasks }
      }
      'bloat' {
        Start-TvTask 'כיבוי אפליקציות' { param($Serial, $P) Invoke-TvDisable $P } @($script:Serial, $pkgs) -Done {
          param($r)
          $u = Read-Undo
          Hide-List
          Write-Log (L 'הסרת זבל' 'head')
          foreach ($x in @($r)) {
            if ($x.Ok) {
              if (-not ($u.Packages | Where-Object { $_.Pkg -eq $x.Pkg })) { $u.Packages += @{ Pkg = $x.Pkg; Method = $x.Method; When = (Get-Date).ToString('s') } }
              Write-Log (L "כובה: $(Get-FriendlyName $x.Pkg)" 'ok')
            } else { Write-Log (L "לא כובה: $($x.Pkg) — $($x.Msg)" 'warn') }
          }
          Save-Undo $u
          Write-Log (L 'כדאי לעשות "הפעלה מחדש" כדי שהזיכרון יתנקה לגמרי.' 'tip')
        }
      }
      'launcher' {
        Start-TvTask 'פתיחת החנות' { param($Serial, $P) Open-TvStore $P } @($script:Serial, $pkgs[0]) -Done { param($r) Hide-List; Write-Log $r }
      }
    }
  })

$ui.ListBtn2.Add_Click({
    switch ($script:ListMode) {
      'tasks' { Open-Tasks }
      'bloat' {
        $sel = Get-Checked
        if (-not $sel.Count) { return }
        $u = Read-Undo
        $entries = @($sel | ForEach-Object { $p = $_.Pkg; $m = @($u.Packages | Where-Object { $_.Pkg -eq $p })[0]; @{ Pkg = $p; Method = $(if ($m) { $m.Method } else { '' }) } })
        Start-TvTask 'החזרת אפליקציות' { param($Serial, $E) Invoke-TvEnable $E } @($script:Serial, $entries) -Done {
          param($r)
          $u = Read-Undo
          $okPkgs = @($r | Where-Object { $_.Ok } | ForEach-Object { $_.Pkg })
          $u.Packages = @($u.Packages | Where-Object { $okPkgs -notcontains $_.Pkg })
          Save-Undo $u
          Write-Log (L "הוחזרו: $(($okPkgs | ForEach-Object { Get-FriendlyName $_ }) -join ', ')" 'ok')
          Open-Bloat
        }
      }
      'launcher' {
        $sel = Get-Checked
        if (-not $sel.Count) { [void][System.Windows.MessageBox]::Show('סמן מסך בית אחד.', 'TVBoost'); return }
        Start-TvTask 'קביעת מסך הבית' { param($Serial, $P) Set-TvHome $P } @($script:Serial, $sel[0].Pkg) -Done {
          param($r)
          $x = @($r)[0]
          Hide-List
          if ($x.Ok) {
            $u = Read-Undo
            if (-not $u.Home -and $x.Prev -and $x.Prev -ne $x.Comp) { $u.Home = $x.Prev; Save-Undo $u }
            Write-Log (L 'מסך הבית הוחלף. תלחץ על כפתור הבית בשלט ותבדוק.' 'ok')
            Write-Log (L 'ב-Projectivy: להחזיק OK על אפליקציה כדי להזיז/להסתיר; להשאיר רק YouTube, Netflix, yes+ ומה שצריך.' 'tip')
            Write-Log (L 'אם כפתור הבית עדיין פותח את המסך הישן: ב-Projectivy ← הגדרות ← "Override current launcher" ולהפעיל.' 'tip')
          } else { Write-Log (L "לא הצלחתי: $($x.Msg)" 'warn') }
        }
      }
    }
  })

$ui.ListBtn3.Add_Click({
    switch ($script:ListMode) {
      'bloat' { Open-Bloat }
      'launcher' {
        $u = Read-Undo
        $target = $u.Home
        if (-not $target) { $target = 'com.google.android.tvlauncher/.MainActivity' }
        Start-TvTask 'החזרת מסך הבית' { param($Serial, $H) Invoke-TvRestore @() @{} $H } @($script:Serial, $target) -Done { param($r) Hide-List; Write-Log $r; $u = Read-Undo; $u.Home = ''; Save-Undo $u }
      }
    }
  })

# ---------------------------------------------------------------- main buttons

$ui.BtnHelp.Add_Click({ Hide-List; Show-AdbHelp })

$ui.BtnScan.Add_Click({
    Hide-List
    Write-Log (L 'מחפש את הטלוויזיה ברשת…' 'head')
    Start-TvTask 'חיפוש ברשת' { Find-TvOnLan } -Done {
      param($r)
      $found = @($r | Where-Object { $_ })
      $adb = @($found | Where-Object { $_.Adb })
      foreach ($f in $adb) { Write-Log (L "נמצאה: $($f.Ip) $($f.Name) — אפשר להתחבר" 'ok') }
      foreach ($f in @($found | Where-Object { -not $_.Adb })) { Write-Log (L "נמצא מכשיר Android TV: $($f.Ip) $($f.Name) — אבל ניפוי באגים (ADB) כבוי בו." 'warn') }
      if ($adb.Count) { $ui.IpBox.Text = $adb[0].Ip; Write-Log (L 'לחץ "התחבר".' 'tip') }
      elseif ($found.Count) { $ui.IpBox.Text = $found[0].Ip; Show-AdbHelp }
      else { Write-Log (L 'לא נמצאה טלוויזיה. ודא שהיא דלוקה (לא בהמתנה) ומחוברת לאותה רשת Wi-Fi.' 'bad'); Show-AdbHelp }
    }
  })

function Start-Connect {
  $ip = $ui.IpBox.Text.Trim()
  if ($ip -notmatch '^\d{1,3}(\.\d{1,3}){3}(:\d+)?$') { Write-Log (L 'כתובת לא תקינה. לחץ "חפש טלוויזיה ברשת" או הקלד כתובת כמו 192.168.1.50' 'warn'); return }
  Hide-List
  Write-Log (L "מתחבר ל-$ip…" 'head')
  Write-Log (L 'אם בטלוויזיה קופץ חלון "לאפשר ניפוי באגים?" — לסמן "אפשר תמיד מהמחשב הזה" וללחוץ OK.' 'tip')
  Start-TvTask 'התחברות (מחכה לאישור בטלוויזיה)' { param($Ip) Connect-Tv $Ip } @($ip) -Done {
    param($r)
    $c = @($r)[0]
    if ($c.Ok) {
      $script:Serial = $c.Target
      $script:DevId = $(if ($c.SerialNo) { $c.SerialNo } else { $c.Target })
      Set-Connected $true "$($c.Model) · Android $($c.Android) · $($c.Build)"
      Save-Settings
      Write-Log (L "מחובר ל-$($c.Model) (Android $($c.Android), $($c.Build))" 'ok')
      Start-Diag 'full'
    } else {
      Set-Connected $false
      if ($c.State -match 'unauthorized') { Write-Log (L 'הטלוויזיה לא אישרה את החיבור. אשר את החלון בטלוויזיה ולחץ שוב "התחבר".' 'warn') }
      else { Write-Log (L "לא הצלחתי להתחבר ($($c.Msg)). ודא שניפוי באגים דלוק והטלוויזיה ערה." 'bad'); Show-AdbHelp }
    }
  }
}
$ui.BtnConnect.Add_Click({ Start-Connect })

function Start-Diag {
  param([string]$Mode)
  Hide-List
  Write-Log (L $(if ($Mode -eq 'net') { 'בדיקת אינטרנט — פינג ומהירות (כ-20 שניות)' } else { 'אבחון מלא — לוקח כחצי דקה' }) 'head')
  Start-TvTask 'אבחון' { param($Serial, $Mode) Get-TvDiagnosis $Mode } @($script:Serial, $Mode) -NeedTv
}
$ui.BtnDiag.Add_Click({ Start-Diag 'full' })
$ui.BtnNet.Add_Click({ Start-Diag 'net' })

$ui.BtnClean.Add_Click({
    Hide-List
    Write-Log (L 'ניקוי זיכרון' 'head')
    Start-TvTask 'ניקוי זיכרון' { param($Serial) Invoke-TvClean } @($script:Serial) -NeedTv
  })

$ui.BtnTasks.Add_Click({ Open-Tasks })
$ui.BtnBloat.Add_Click({ Open-Bloat })
$ui.BtnLauncher.Add_Click({ Open-Launchers })

$ui.BtnTweaks.Add_Click({
    Hide-List
    Write-Log (L 'הגדרות מהירות' 'head')
    Start-TvTask 'הגדרות מהירות' { param($Serial) Invoke-TvTweaks } @($script:Serial) -NeedTv -Done {
      param($r)
      $x = @($r)[0]
      $u = Read-Undo
      foreach ($k in $x.Old.Keys) { if (-not $u.Settings.ContainsKey($k)) { $u.Settings[$k] = $x.Old[$k] } }
      Save-Undo $u
      Write-Log (L 'אנימציות קוצרו לחצי — התפריטים ירגישו מהירים יותר.' 'ok')
      Write-Log (L 'כובו סריקות Wi-Fi ובלוטות'' ברקע (חוסך מעבד ומפריע פחות לסטרימינג).' 'ok')
    }
  })

$ui.BtnReboot.Add_Click({
    if (-not $script:Serial) { Write-Log (L 'קודם צריך להתחבר לטלוויזיה.' 'warn'); return }
    if ([System.Windows.MessageBox]::Show('להפעיל מחדש את הטלוויזיה עכשיו?', 'TVBoost', 'YesNo') -ne 'Yes') { return }
    Hide-List
    Start-TvTask 'הפעלה מחדש' { param($Serial) [void](Invoke-Adb @('-s', $Serial, 'reboot') 15); L 'הטלוויזיה עושה הפעלה מחדש. כשהיא עולה (כדקה) — לחץ "התחבר".' 'ok' } @($script:Serial) -NeedTv -Done {
      param($r); Write-Log $r; Set-Connected $false
    }
  })

$ui.BtnRestore.Add_Click({
    if (-not $script:Serial) { Write-Log (L 'קודם צריך להתחבר לטלוויזיה.' 'warn'); return }
    $u = Read-Undo
    if (-not $u.Packages.Count -and -not $u.Settings.Count -and -not $u.Home) { Hide-List; Write-Log (L 'אין מה לשחזר — הכלי עוד לא שינה כלום בטלוויזיה הזו.' 'ok'); return }
    if ([System.Windows.MessageBox]::Show("להחזיר את כל השינויים? ($($u.Packages.Count) אפליקציות, $($u.Settings.Count) הגדרות)", 'TVBoost', 'YesNo') -ne 'Yes') { return }
    Hide-List
    Write-Log (L 'שחזור' 'head')
    Start-TvTask 'שחזור' { param($Serial, $P, $S, $H) Invoke-TvRestore $P $S $H } @($script:Serial, @($u.Packages), $u.Settings, $u.Home) -Done {
      param($r)
      $u = Read-Undo
      foreach ($x in @($r)) { if ($x.K -eq 'restored') { $u.Packages = @($u.Packages | Where-Object { $x.Pkgs -notcontains $_.Pkg }) } }
      $u.Settings = @{}; $u.Home = ''
      Save-Undo $u
      Write-Log $r
    }
  })

$ui.BtnSave.Add_Click({
    $path = Join-Path $ReportDir ("TVBoost-" + (Get-Date -Format 'yyyy-MM-dd_HH-mm') + '.txt')
    ($script:LogText -join "`r`n") | Set-Content -Path $path -Encoding UTF8
    Write-Log (L "הדוח נשמר: $path" 'ok')
    Start-Process explorer.exe "/select,`"$path`""
  })

# ---------------------------------------------------------------- start

if ($SelfTest) {
  $fails = Invoke-SelfTest
  if ($RenderPng) {
    Write-Log @((L 'ברוך הבא ל-TVBoost' 'head'), (L 'דוגמה לשורה תקינה' 'ok'), (L 'דוגמה לאזהרה' 'warn'), (L 'דוגמה לבעיה' 'bad'), (L 'דוגמה להמלצה' 'tip'), (L '   181 MB  מערכת Android' 'dim'))
    Set-Connected $true 'MiTV-MSSP3 · Android 9 · PI.4663'
    $Win.Width = 1200; $Win.Height = 790
    $Win.WindowStartupLocation = 'Manual'; $Win.Left = -3000; $Win.Top = 0; $Win.ShowInTaskbar = $false
    $Win.Show()
    $Win.Dispatcher.Invoke([action] {}, 'Render')
    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap 1200, 790, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($Win.Content)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $fs = [IO.File]::Create($RenderPng); $enc.Save($fs); $fs.Close()
    $Win.Close()
    Write-Host "rendered $RenderPng"
  }
  exit $fails
}

$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
foreach ($fn in $BgFunctions) {
  $iss.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $fn, (Get-Item "function:$fn").ScriptBlock.ToString()))
}
foreach ($v in $BgVariables) {
  $iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $v, (Get-Variable $v -ValueOnly), ''))
}
$script:Pool = [runspacefactory]::CreateRunspacePool(1, 1, $iss, $Host)
$script:Pool.Open()
$Win.Add_Closed({ $Timer.Stop(); try { $script:Pool.Close() } catch { } })

$ui.FootText.Text = $(if ($AdbPath) { "adb: $AdbPath" } else { 'adb לא נמצא' })
Write-Log @(
  (L 'ברוך הבא ל-TVBoost' 'head')
  (L 'הכלי רץ על המחשב ומתחבר לטלוויזיה דרך ה-Wi-Fi, כך שהוא לא תופס לה זיכרון. הוא עושה אבחון (זיכרון, מעבד, אחסון, אינטרנט), ניקוי זיכרון, מנהל משימות, כיבוי אפליקציות זבל, מסך בית נקי ובדיקת מהירות Wi-Fi.')
  (L 'שום דבר לא נמחק מהטלוויזיה. כל שינוי נרשם וחוזר בכפתור "שחזור הכל".' 'dim')
)
if (-not $AdbPath) { Write-Log (L 'לא נמצא adb.exe במחשב. צריך להתקין Android SDK Platform-Tools מהאתר של גוגל, או לשים את התיקייה platform-tools ליד הקובץ הזה.' 'bad') }
else {
  # בלי -Wait: הוא מחכה גם לשרת ה-adb שנשאר לרוץ ברקע
  $sp = Start-Process -FilePath $AdbPath -ArgumentList 'start-server' -WindowStyle Hidden -PassThru
  [void]$sp.WaitForExit(15000)
}

$settingsPath = Join-Path $DataDir 'settings.json'
$lastIp = ''
if (Test-Path $settingsPath) { try { $lastIp = (Get-Content $settingsPath -Raw | ConvertFrom-Json).LastIp } catch { } }
if ($lastIp) {
  $ui.IpBox.Text = $lastIp
  $tcp = New-Object System.Net.Sockets.TcpClient
  $open = $false
  try { $open = $tcp.ConnectAsync($lastIp, 5555).Wait(700) -and $tcp.Connected } catch { }
  $tcp.Dispose()
  if ($open -and $AdbPath) { Start-Connect } else { Write-Log (L "כתובת אחרונה: $lastIp — הטלוויזיה לא עונה כרגע. כשהיא דלוקה לחץ ""התחבר""." 'dim') }
} else { Show-AdbHelp }

[void]$Win.ShowDialog()
