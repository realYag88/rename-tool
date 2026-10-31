# =============================================================
#  リネームツール (RenameTool.ps1)
#  連番リネーム / 文字列置換リネーム
#  Windows PowerShell 5.1 + WinForms（追加インストール不要）
# =============================================================
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptPath   = $MyInvocation.MyCommand.Path
$script:InitialPaths = @($args)
$script:AppTitle     = 'リネームツール'

# ---------- 「送る」メニューへの登録 ----------
if ($script:InitialPaths.Count -gt 0 -and $script:InitialPaths[0] -eq '--install-sendto') {
    $shell   = New-Object -ComObject WScript.Shell
    $lnkPath = Join-Path ([Environment]::GetFolderPath('SendTo')) 'リネームツール.lnk'
    $lnk     = $shell.CreateShortcut($lnkPath)
    $lnk.TargetPath  = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $lnk.Arguments   = '-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "' + $script:ScriptPath + '"'
    $lnk.WindowStyle = 7
    $lnk.Save()
    [System.Windows.Forms.MessageBox]::Show(
        "「送る」メニューに登録しました。`n`nエクスプローラーでファイルを選択 → 右クリック →`n（Win11は「その他のオプションを確認」）→ 送る → リネームツール",
        $script:AppTitle) | Out-Null
    exit
}

[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------- 状態 ----------
$script:Items   = New-Object 'System.Collections.Generic.List[object]'
$script:PathSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$script:UndoLog = $null
$script:InvalidChars = [IO.Path]::GetInvalidFileNameChars()

function Show-Error([string]$msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, $script:AppTitle,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}

function Safe([scriptblock]$sb) {
    try { & $sb } catch { Show-Error ("エラー: " + $_.Exception.Message) }
}

function New-Entry([string]$path) {
    [pscustomobject]@{
        Path    = $path
        Dir     = [IO.Path]::GetDirectoryName($path)
        Name    = [IO.Path]::GetFileName($path)
        NewName = ''
        Final   = ''
        Changed = $false
        IsError = $false
        Status  = ''
    }
}

# 数字を桁埋めして「自然順」で並べるためのキー（frame2 < frame10）
function Get-NaturalKey([string]$s) {
    [regex]::Replace($s.ToLowerInvariant(), '\d+', { param($m) $m.Value.PadLeft(20, '0') })
}

function Get-ExtFilter {
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($e in ($txtFilter.Text -split '[,;\s]+')) {
        $e = $e.Trim().TrimStart('*').TrimStart('.')
        if ($e -ne '') { [void]$set.Add($e) }
    }
    return ,$set
}

# =============================================================
#  リネーム本体（一時名を経由する2段階方式。失敗したら全部戻す）
# =============================================================
function Invoke-RenameBatch($pairs) {
    try {
        foreach ($p in $pairs) {
            $tmp = [IO.Path]::Combine([IO.Path]::GetDirectoryName($p.From),
                   '~rn_' + [guid]::NewGuid().ToString('N') + '.tmp')
            [IO.File]::Move($p.From, $tmp)
            $p.Current = $tmp
        }
        foreach ($p in $pairs) {
            [IO.File]::Move($p.Current, $p.To)
            $p.Current = $p.To
        }
        return $null
    }
    catch {
        $err = $_.Exception.Message
        # ロールバック：入れ替え同士で詰まらないよう、進まなくなるまで繰り返す
        $pending = New-Object System.Collections.ArrayList
        foreach ($p in $pairs) { if ($p.Current -and $p.Current -ne $p.From) { [void]$pending.Add($p) } }
        do {
            $progress = $false
            foreach ($p in @($pending)) {
                try {
                    [IO.File]::Move($p.Current, $p.From)
                    $p.Current = $p.From
                    $pending.Remove($p)
                    $progress = $true
                } catch { }
            }
        } while ($progress -and $pending.Count -gt 0)
        if ($pending.Count -gt 0) {
            $left = ($pending | ForEach-Object { $_.Current }) -join "`n"
            $err += "`n`n一部のファイルを元に戻せませんでした:`n" + $left
        }
        return $err
    }
}

# =============================================================
#  画面
# =============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = $script:AppTitle
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
$form.Font = New-Object System.Drawing.Font('Yu Gothic UI', 9)
$form.ClientSize = New-Object System.Drawing.Size(944, 640)
$form.MinimumSize = New-Object System.Drawing.Size(820, 560)
$form.StartPosition = 'CenterScreen'
$form.AllowDrop = $true

function New-Button($text, $x, $y, $w, $h, $anchor) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.SetBounds($x, $y, $w, $h)
    $b.Anchor = $anchor
    $b
}
function New-Label($text, $x, $y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.AutoSize = $true
    $l.Location = New-Object System.Drawing.Point($x, $y)
    $l
}
$TL  = [System.Windows.Forms.AnchorStyles]'Top,Left'
$TR  = [System.Windows.Forms.AnchorStyles]'Top,Right'
$BL  = [System.Windows.Forms.AnchorStyles]'Bottom,Left'
$BR  = [System.Windows.Forms.AnchorStyles]'Bottom,Right'
$BLR = [System.Windows.Forms.AnchorStyles]'Bottom,Left,Right'
$ALL = [System.Windows.Forms.AnchorStyles]'Top,Bottom,Left,Right'

# --- 上段：追加・並べ替え ---
$btnAddFolder = New-Button 'フォルダを追加' 10 10 110 28 $TL
$btnAddFiles  = New-Button 'ファイルを追加' 125 10 110 28 $TL
$chkRecurse = New-Object System.Windows.Forms.CheckBox
$chkRecurse.Text = 'サブフォルダも'
$chkRecurse.AutoSize = $true
$chkRecurse.Location = New-Object System.Drawing.Point(245, 15)
$lblFilter = New-Label '拡張子:' 360 16
$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.SetBounds(410, 13, 90, 24)
$lblFilterHint = New-Label '例 png,exr' 505 16
$lblFilterHint.ForeColor = [System.Drawing.Color]::Gray

$btnUp     = New-Button '↑' 594 10 34 28 $TR
$btnDown   = New-Button '↓' 632 10 34 28 $TR
$btnSort   = New-Button '名前順' 672 10 70 28 $TR
$btnRemove = New-Button 'リストから外す' 746 10 104 28 $TR
$btnClear  = New-Button 'クリア' 854 10 80 28 $TR

# --- リスト ---
$lv = New-Object System.Windows.Forms.ListView
$lv.SetBounds(10, 48, 924, 380)
$lv.Anchor = $ALL
$lv.View = 'Details'
$lv.FullRowSelect = $true
$lv.GridLines = $true
$lv.HideSelection = $false
$lv.MultiSelect = $true
$lv.AllowDrop = $true
[void]$lv.Columns.Add('#', 50)
[void]$lv.Columns.Add('元の名前', 250)
[void]$lv.Columns.Add('新しい名前', 250)
[void]$lv.Columns.Add('状態', 100)
[void]$lv.Columns.Add('フォルダ', 250)

# --- 対象 ---
$rbAll = New-Object System.Windows.Forms.RadioButton
$rbAll.Text = 'リスト全体'
$rbAll.AutoSize = $true
$rbAll.Checked = $true
$rbAll.Location = New-Object System.Drawing.Point(12, 436)
$rbAll.Anchor = $BL
$rbSel = New-Object System.Windows.Forms.RadioButton
$rbSel.Text = '選択した行だけ'
$rbSel.AutoSize = $true
$rbSel.Location = New-Object System.Drawing.Point(110, 436)
$rbSel.Anchor = $BL
$lblHint = New-Label 'ドラッグ＆ドロップでも追加できます。連番はリストの並び順で振られます。' 240 438
$lblHint.ForeColor = [System.Drawing.Color]::Gray
$lblHint.Anchor = $BL
$lnkSite = New-Object System.Windows.Forms.LinkLabel
$lnkSite.Text = 'caprame.com'
$lnkSite.AutoSize = $true
$lnkSite.Location = New-Object System.Drawing.Point(850, 438)
$lnkSite.Anchor = $BR
$lnkSite.Add_LinkClicked({ Start-Process 'https://caprame.com/tools/rename-tool/' })

# --- モード（タブ） ---
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.SetBounds(10, 462, 924, 128)
$tabs.Anchor = $BLR
$tabSeq = New-Object System.Windows.Forms.TabPage
$tabSeq.Text = '  連番  '
$tabRep = New-Object System.Windows.Forms.TabPage
$tabRep.Text = '  置換  '
$tabs.TabPages.Add($tabSeq)
$tabs.TabPages.Add($tabRep)

function New-Num($x, $y, $w, $min, $max, $val) {
    $n = New-Object System.Windows.Forms.NumericUpDown
    $n.SetBounds($x, $y, $w, 24)
    $n.Minimum = $min
    $n.Maximum = $max
    $n.Value = $val
    $n
}

# 連番タブ
$txtBase = New-Object System.Windows.Forms.TextBox
$txtBase.SetBounds(80, 16, 230, 24)
$txtSep = New-Object System.Windows.Forms.TextBox
$txtSep.SetBounds(372, 16, 46, 24)
$txtSep.Text = '_'
$numStart  = New-Num 470 16 80 0 99999999 1
$numDigits = New-Num 610 16 54 1 12 4
$numStep   = New-Num 722 16 64 1 10000 1
$lblSeqEx = New-Label '' 14 56
$lblSeqEx.ForeColor = [System.Drawing.Color]::FromArgb(0, 90, 180)
$tabSeq.Controls.AddRange(@(
    (New-Label 'ベース名' 14 19), $txtBase,
    (New-Label '区切り' 324 19), $txtSep,
    (New-Label '開始' 432 19), $numStart,
    (New-Label '桁数' 572 19), $numDigits,
    (New-Label '増分' 684 19), $numStep,
    $lblSeqEx))

# 置換タブ
$txtFind = New-Object System.Windows.Forms.TextBox
$txtFind.SetBounds(60, 16, 280, 24)
$txtRep = New-Object System.Windows.Forms.TextBox
$txtRep.SetBounds(410, 16, 280, 24)
$chkRegex = New-Object System.Windows.Forms.CheckBox
$chkRegex.Text = '正規表現'
$chkRegex.AutoSize = $true
$chkRegex.Location = New-Object System.Drawing.Point(16, 54)
$chkCase = New-Object System.Windows.Forms.CheckBox
$chkCase.Text = '大文字/小文字を区別'
$chkCase.AutoSize = $true
$chkCase.Location = New-Object System.Drawing.Point(110, 54)
$chkExt = New-Object System.Windows.Forms.CheckBox
$chkExt.Text = '拡張子も対象にする'
$chkExt.AutoSize = $true
$chkExt.Location = New-Object System.Drawing.Point(270, 54)
$tabRep.Controls.AddRange(@(
    (New-Label '検索' 14 19), $txtFind,
    (New-Label '置換後' 354 19), $txtRep,
    $chkRegex, $chkCase, $chkExt))

# --- 下段 ---
$lblStatus = New-Label '' 12 607
$lblStatus.Anchor = $BL
$btnUndo = New-Button '元に戻す' 700 598 110 34 $BR
$btnUndo.Enabled = $false
$btnRun  = New-Button 'リネーム実行' 820 598 114 34 $BR
$btnRun.Font = New-Object System.Drawing.Font('Yu Gothic UI', 9, [System.Drawing.FontStyle]::Bold)
$btnRun.Enabled = $false

$form.Controls.AddRange(@(
    $btnAddFolder, $btnAddFiles, $chkRecurse, $lblFilter, $txtFilter, $lblFilterHint,
    $btnUp, $btnDown, $btnSort, $btnRemove, $btnClear,
    $lv, $rbAll, $rbSel, $lblHint, $lnkSite, $tabs, $lblStatus, $btnUndo, $btnRun))

# =============================================================
#  リスト操作
# =============================================================
function Rebuild-PathSet {
    $script:PathSet.Clear()
    foreach ($it in $script:Items) { [void]$script:PathSet.Add($it.Path) }
}

function Refresh-List {
    $lv.BeginUpdate()
    $lv.Items.Clear()
    $rows = New-Object 'System.Collections.Generic.List[System.Windows.Forms.ListViewItem]'
    $i = 1
    foreach ($it in $script:Items) {
        $row = New-Object System.Windows.Forms.ListViewItem -ArgumentList ([string]$i)
        $row.UseItemStyleForSubItems = $false
        [void]$row.SubItems.Add($it.Name)
        [void]$row.SubItems.Add('')
        [void]$row.SubItems.Add('')
        [void]$row.SubItems.Add($it.Dir)
        $rows.Add($row)
        $i++
    }
    $lv.Items.AddRange($rows.ToArray())
    $lv.EndUpdate()
    Update-Preview
}

function Add-Paths([string[]]$paths) {
    $filter  = Get-ExtFilter
    $recurse = $chkRecurse.Checked
    $batch = New-Object 'System.Collections.Generic.List[string]'
    foreach ($p in $paths) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if (Test-Path -LiteralPath $p -PathType Container) {
            if ($recurse) { $files = Get-ChildItem -LiteralPath $p -File -Recurse -ErrorAction SilentlyContinue }
            else          { $files = Get-ChildItem -LiteralPath $p -File -ErrorAction SilentlyContinue }
            foreach ($f in $files) {
                if ($filter.Count -gt 0 -and -not $filter.Contains($f.Extension.TrimStart('.'))) { continue }
                $batch.Add($f.FullName)
            }
        }
        elseif (Test-Path -LiteralPath $p -PathType Leaf) {
            $batch.Add((Get-Item -LiteralPath $p -Force).FullName)
        }
    }
    $sorted = $batch | Sort-Object @{ Expression = { [IO.Path]::GetDirectoryName($_).ToLowerInvariant() } },
                                   @{ Expression = { Get-NaturalKey ([IO.Path]::GetFileName($_)) } }
    foreach ($s in $sorted) {
        if ($script:PathSet.Add($s)) { $script:Items.Add((New-Entry $s)) }
    }
    # ベース名が空なら、先頭ファイル名から末尾の番号を除いたものを入れておく
    if ($txtBase.Text -eq '' -and $script:Items.Count -gt 0) {
        $stem = [IO.Path]::GetFileNameWithoutExtension($script:Items[0].Name)
        $txtBase.Text = [regex]::Replace($stem, '[_\-\. ]*\d+$', '')
    }
    Refresh-List
}

function Get-SelectedSorted {
    return @($lv.SelectedIndices | ForEach-Object { [int]$_ } | Sort-Object)
}

function Select-Rows([int[]]$idx) {
    foreach ($i in $idx) { if ($i -ge 0 -and $i -lt $lv.Items.Count) { $lv.Items[$i].Selected = $true } }
    if ($idx.Count -gt 0) { $lv.EnsureVisible($idx[0]) }
}

function Move-Selected([int]$dir) {
    $sel = Get-SelectedSorted
    if ($sel.Count -eq 0) { return }
    if ($dir -lt 0) {
        if ($sel[0] -eq 0) { return }
        foreach ($i in $sel) {
            $t = $script:Items[$i - 1]; $script:Items[$i - 1] = $script:Items[$i]; $script:Items[$i] = $t
        }
    } else {
        if ($sel[-1] -ge $script:Items.Count - 1) { return }
        foreach ($i in ($sel | Sort-Object -Descending)) {
            $t = $script:Items[$i + 1]; $script:Items[$i + 1] = $script:Items[$i]; $script:Items[$i] = $t
        }
    }
    Refresh-List
    Select-Rows @($sel | ForEach-Object { $_ + $dir })
}

function Remove-Selected {
    $sel = @(Get-SelectedSorted | Sort-Object -Descending)
    if ($sel.Count -eq 0) { return }
    foreach ($i in $sel) { $script:Items.RemoveAt($i) }
    Rebuild-PathSet
    Refresh-List
}

function Sort-All {
    $sorted = @($script:Items | Sort-Object @{ Expression = { $_.Dir.ToLowerInvariant() } },
                                            @{ Expression = { Get-NaturalKey $_.Name } })
    $script:Items.Clear()
    foreach ($s in $sorted) { $script:Items.Add($s) }
    Refresh-List
}

# =============================================================
#  プレビュー（新しい名前の計算とチェック）
# =============================================================
function Get-TargetIndices {
    if ($script:Items.Count -eq 0) { return @() }
    if ($rbSel.Checked) { return @(Get-SelectedSorted) }
    return @(0..($script:Items.Count - 1))
}

function Update-Preview {
    foreach ($it in $script:Items) { $it.NewName = $it.Name }
    $targets = @(Get-TargetIndices)
    $regexError = $null

    if ($tabs.SelectedIndex -eq 0) {
        # ---- 連番 ----
        $base   = $txtBase.Text
        $sep    = $txtSep.Text
        $n      = [long]$numStart.Value
        $digits = [int]$numDigits.Value
        $step   = [long]$numStep.Value
        foreach ($idx in $targets) {
            $it  = $script:Items[$idx]
            $num = ([string]$n).PadLeft($digits, '0')
            $ext = [IO.Path]::GetExtension($it.Name)
            if ($base -eq '') { $it.NewName = $num + $ext }
            else              { $it.NewName = $base + $sep + $num + $ext }
            $n += $step
        }
        if ($targets.Count -gt 0) {
            $lblSeqEx.Text = '→ ' + $script:Items[$targets[0]].NewName + '  〜  ' + $script:Items[$targets[-1]].NewName + '（拡張子はそのまま）'
        } else {
            $ex = ([string]$numStart.Value).PadLeft($digits, '0')
            if ($base -ne '') { $ex = $base + $sep + $ex }
            $lblSeqEx.Text = '→ 例: ' + $ex + '.png（拡張子はそのまま）'
        }
    }
    else {
        # ---- 置換 ----
        $find = $txtFind.Text
        $rep  = $txtRep.Text
        if ($find -ne '') {
            $rx = $null
            try {
                $opts = [System.Text.RegularExpressions.RegexOptions]::None
                if (-not $chkCase.Checked) { $opts = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
                if ($chkRegex.Checked) { $rx = [regex]::new($find, $opts); $repl = $rep }
                else                   { $rx = [regex]::new([regex]::Escape($find), $opts); $repl = $rep.Replace('$', '$$') }
            } catch { $regexError = $_.Exception.Message }
            if ($rx) {
                foreach ($idx in $targets) {
                    $it = $script:Items[$idx]
                    if ($chkExt.Checked) {
                        $it.NewName = $rx.Replace($it.Name, $repl)
                    } else {
                        $ext  = [IO.Path]::GetExtension($it.Name)
                        $stem = $it.Name.Substring(0, $it.Name.Length - $ext.Length)
                        $it.NewName = $rx.Replace($stem, $repl) + $ext
                    }
                }
            }
        }
    }

    # ---- チェック ----
    $counts  = New-Object 'System.Collections.Generic.Dictionary[string,int]' ([StringComparer]::OrdinalIgnoreCase)
    $vacated = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($it in $script:Items) {
        $it.Changed = ($it.NewName -cne $it.Name)
        $it.IsError = $false
        $it.Final   = $it.Dir.TrimEnd('\') + '\' + $it.NewName
        if ($it.Changed) { [void]$vacated.Add($it.Path) }
        if ($counts.ContainsKey($it.Final)) { $counts[$it.Final] = $counts[$it.Final] + 1 }
        else { $counts[$it.Final] = 1 }
    }
    $nChanged = 0; $nErr = 0
    foreach ($it in $script:Items) {
        $nm = $it.NewName
        if ($nm.IndexOfAny($script:InvalidChars) -ge 0) { $it.Status = '使えない文字'; $it.IsError = $true }
        elseif ($nm.Trim() -eq '' -or $nm.EndsWith('.') -or $nm.EndsWith(' ')) { $it.Status = '名前が不正'; $it.IsError = $true }
        elseif ($counts[$it.Final] -gt 1) { $it.Status = '名前が重複'; $it.IsError = $true }
        elseif ($it.Changed -and -not $vacated.Contains($it.Final) -and
                ([IO.File]::Exists($it.Final) -or [IO.Directory]::Exists($it.Final))) {
            $it.Status = '既存と衝突'; $it.IsError = $true
        }
        elseif ($it.Changed) { $it.Status = 'OK' }
        else { $it.Status = '変更なし' }
        if ($it.IsError) { $nErr++ } elseif ($it.Changed) { $nChanged++ }
    }

    # ---- 表示 ----
    if ($lv.Items.Count -eq $script:Items.Count) {
        $red   = [System.Drawing.Color]::Firebrick
        $blue  = [System.Drawing.Color]::FromArgb(0, 90, 180)
        $green = [System.Drawing.Color]::FromArgb(0, 130, 60)
        $gray  = [System.Drawing.Color]::Gray
        $lv.BeginUpdate()
        for ($i = 0; $i -lt $script:Items.Count; $i++) {
            $it  = $script:Items[$i]
            $row = $lv.Items[$i]
            $row.SubItems[2].Text = $it.NewName
            $row.SubItems[3].Text = $it.Status
            if     ($it.IsError) { $c2 = $red;  $c3 = $red }
            elseif ($it.Changed) { $c2 = $blue; $c3 = $green }
            else                 { $c2 = $gray; $c3 = $gray }
            $row.SubItems[2].ForeColor = $c2
            $row.SubItems[3].ForeColor = $c3
        }
        $lv.EndUpdate()
    }

    if ($regexError) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text = '正規表現エラー: ' + $regexError
    } else {
        $msg = '{0} 件 ／ 変更 {1} 件' -f $script:Items.Count, $nChanged
        if ($nErr -gt 0) {
            $msg += ' ／ エラー {0} 件（解消するまで実行できません）' -f $nErr
            $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        } else {
            $lblStatus.ForeColor = [System.Drawing.SystemColors]::ControlText
        }
        $lblStatus.Text = $msg
    }
    $btnRun.Enabled = ($nChanged -gt 0 -and $nErr -eq 0 -and -not $regexError)
}

# 入力のたびに即計算すると重いので少し待ってから更新
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 120
$timer.Add_Tick({ $timer.Stop(); Safe { Update-Preview } })
function Request-Preview { $timer.Stop(); $timer.Start() }

# =============================================================
#  実行・元に戻す
# =============================================================
function Run-Rename {
    Update-Preview
    if (-not $btnRun.Enabled) { return }
    $pairs = New-Object System.Collections.ArrayList
    foreach ($it in $script:Items) {
        if ($it.Changed) {
            [void]$pairs.Add([pscustomobject]@{ Entry = $it; From = $it.Path; To = $it.Final; Current = $null })
        }
    }
    $err = Invoke-RenameBatch $pairs
    if ($err) {
        Show-Error ("リネームに失敗したため、元の名前に戻しました。`n（ファイルが他のソフトで開かれていないか確認してください）`n`n" + $err)
        Update-Preview
        return
    }
    foreach ($p in $pairs) {
        $p.Entry.Path = $p.To
        $p.Entry.Name = [IO.Path]::GetFileName($p.To)
    }
    $script:UndoLog = $pairs
    $btnUndo.Enabled = $true
    Rebuild-PathSet
    Refresh-List
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 130, 60)
    $lblStatus.Text = ('{0} 件リネームしました。' -f $pairs.Count) + '  ' + $lblStatus.Text
}

function Run-Undo {
    if (-not $script:UndoLog) { return }
    $pairs = New-Object System.Collections.ArrayList
    foreach ($u in $script:UndoLog) {
        [void]$pairs.Add([pscustomobject]@{ Entry = $u.Entry; From = $u.To; To = $u.From; Current = $null })
    }
    $err = Invoke-RenameBatch $pairs
    if ($err) {
        Show-Error ("元に戻せませんでした（変更はしていません）。`n`n" + $err)
        return
    }
    foreach ($p in $pairs) {
        $p.Entry.Path = $p.To
        $p.Entry.Name = [IO.Path]::GetFileName($p.To)
    }
    $script:UndoLog = $null
    $btnUndo.Enabled = $false
    Rebuild-PathSet
    Refresh-List
    $lblStatus.Text = ('{0} 件を元に戻しました。' -f $pairs.Count) + '  ' + $lblStatus.Text
}

# =============================================================
#  イベント
# =============================================================
$btnAddFolder.Add_Click({ Safe {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'リネームするファイルが入ったフォルダを選択'
    if ($script:LastFolder) { $dlg.SelectedPath = $script:LastFolder }
    if ($dlg.ShowDialog($form) -eq 'OK') {
        $script:LastFolder = $dlg.SelectedPath
        Add-Paths @($dlg.SelectedPath)
    }
} })

$btnAddFiles.Add_Click({ Safe {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Multiselect = $true
    $dlg.Title = 'リネームするファイルを選択'
    $dlg.Filter = 'すべてのファイル (*.*)|*.*'
    if ($script:LastFolder) { $dlg.InitialDirectory = $script:LastFolder }
    if ($dlg.ShowDialog($form) -eq 'OK') {
        $script:LastFolder = [IO.Path]::GetDirectoryName($dlg.FileNames[0])
        Add-Paths $dlg.FileNames
    }
} })

$btnUp.Add_Click({ Safe { Move-Selected -1 } })
$btnDown.Add_Click({ Safe { Move-Selected 1 } })
$btnSort.Add_Click({ Safe { Sort-All } })
$btnRemove.Add_Click({ Safe { Remove-Selected } })
$btnClear.Add_Click({ Safe {
    $script:Items.Clear()
    Rebuild-PathSet
    Refresh-List
} })
$btnRun.Add_Click({ Safe { Run-Rename } })
$btnUndo.Add_Click({ Safe { Run-Undo } })

$dragEnter = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
}
$dragDrop = {
    param($s, $e)
    $files = [string[]]$e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    Safe { Add-Paths $files }
}
$lv.Add_DragEnter($dragEnter)
$lv.Add_DragDrop($dragDrop)
$form.Add_DragEnter($dragEnter)
$form.Add_DragDrop($dragDrop)

$lv.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) { Safe { Remove-Selected } }
    elseif ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
        $lv.BeginUpdate()
        foreach ($row in $lv.Items) { $row.Selected = $true }
        $lv.EndUpdate()
    }
})
$lv.Add_SelectedIndexChanged({ if ($rbSel.Checked) { Request-Preview } })

foreach ($c in @($txtBase, $txtSep, $txtFind, $txtRep)) { $c.Add_TextChanged({ Request-Preview }) }
foreach ($c in @($numStart, $numDigits, $numStep))       { $c.Add_ValueChanged({ Request-Preview }) }
foreach ($c in @($chkRegex, $chkCase, $chkExt, $rbAll, $rbSel)) { $c.Add_CheckedChanged({ Request-Preview }) }
$tabs.Add_SelectedIndexChanged({ Request-Preview })

$form.Add_Shown({
    Safe {
        if ($script:InitialPaths.Count -gt 0) { Add-Paths $script:InitialPaths }
        else { Update-Preview }
    }
    $form.Activate()
})

try {
    [void]$form.ShowDialog()
} catch {
    Show-Error ("起動中にエラーが発生しました:`n" + $_.Exception.Message)
}
