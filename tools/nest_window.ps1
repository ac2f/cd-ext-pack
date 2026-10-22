param([string]$In, [string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$inv = [System.Globalization.CultureInfo]::InvariantCulture
$all = [System.IO.File]::ReadAllLines($In, [System.Text.Encoding]::UTF8)
$vals = @{}
$sheets = New-Object System.Collections.Generic.List[object]
$presets = New-Object System.Collections.Specialized.OrderedDictionary
$title = 'Nesting'
$summary = ''
$pageW = ''
$pageH = ''
foreach ($ln in $all) {
  $p = $ln.Split('|')
  switch ($p[0]) {
    'T' { $title = $p[1] }
    'P' { $summary = $p[1] }
    'G' { if ($p.Length -ge 3) { $pageW = $p[1]; $pageH = $p[2] } }
    'K' { $vals[$p[1]] = $p[2] }
    'S' { if ($p.Length -ge 7) { $sheets.Add(@($p[1], $p[2], $p[3], $p[4], $p[5], $p[6])) } }
    'R' { if ($p.Length -ge 3) { $presets[$p[1]] = $p[2] } }
  }
}

# key, label, kind, options, tab, help
$fields = @(
  @('gap', 'Part gap (mm)', 'num', '', 0, 'Clear space between neighbouring parts. A router needs at least its bit diameter, a laser its kerf plus a little.'),
  @('mL', 'Margin left (mm)', 'num', '', 0, 'Clear strip along the left edge of every sheet.'),
  @('mR', 'Margin right (mm)', 'num', '', 0, 'Clear strip along the right edge of every sheet.'),
  @('mT', 'Margin top (mm)', 'num', '', 0, 'Clear strip along the top edge of every sheet.'),
  @('mB', 'Margin bottom (mm)', 'num', '', 0, 'Clear strip along the bottom edge of every sheet.'),
  @('rot', 'Rotation', 'combo', '0=Keep as drawn|180=Half turn only|90=Every 90 degrees|45=Every 45 degrees|30=Every 30 degrees|15=Every 15 degrees', 0, 'The turns parts may take. Finer steps nest tighter but take longer.'),
  @('copies', 'Copies of each part', 'int', '', 0, 'Every selected part is nested this many times. Copies are only created when they find room.'),
  @('time', 'Optimise for (seconds)', 'num', '', 0, 'Keeps trying other orders and strategies this long and keeps the best layout. 0 = a single quick pass.'),
  @('inner', 'Shapes inside a shape move with it', 'chk', '', 0, 'A shape lying inside another (a counter, a hole, an engraving) is kept with the outer shape as one part.'),
  @('holes', 'Small parts may go into holes', 'chk', '', 0, 'Inner contours are cut out, so other parts may be nested inside them. Leave off when inner shapes are engravings.'),
  @('om', 'Offcut edge (mm)', 'num', '', 1, 'Clear space kept from the edge of a drawn offcut sheet.'),
  @('mirror', 'Parts may be mirrored', 'chk', '', 1, 'Allows mirrored copies. Leave off for printed, coated or filmed material.'),
  @('corner', 'Start corner', 'combo', '1=Bottom left|2=Bottom right|3=Top left|4=Top right', 1, 'The corner of the sheet the packing starts from.'),
  @('dir', 'Fill direction', 'combo', '0=Rows (fill along the width)|1=Columns (fill along the height)', 1, 'Rows keep the used material short in height; columns keep it short in width.'),
  @('strat', 'Placement', 'combo', '0=Automatic (tries both)|1=Lowest top edge|2=Bottom-left', 1, 'How each part chooses its place. Automatic tries both and keeps the better layout.'),
  @('res', 'Precision (mm)', 'combo', '0.5=0.5 mm (finest, slowest)|1=1 mm|2=2 mm|3=3 mm|5=5 mm (fastest)', 1, 'Grid size used to test shapes. Finer is tighter and slower. The gap is always kept.'),
  @('group', 'Group the result', 'combo', '0=No grouping|1=Each part|2=Each sheet|3=Everything', 1, 'How placed parts are grouped once nesting is done.'),
  @('outline', 'Draw sheet outlines', 'chk', '', 1, 'Draws a grey outline for every sheet used.'),
  @('start', 'Set start points afterwards', 'chk', '', 1, 'Applies the saved start point settings (main menu) to every placed part.'),
  @('sgap', 'Space between sheets (mm)', 'num', '', 1, 'Gap between the sheets drawn on the page. Layout only.')
)

$f = New-Object System.Windows.Forms.Form
$f.Text = $title
$f.StartPosition = 'CenterScreen'
$f.Size = New-Object System.Drawing.Size(1060, 680)
$f.MinimumSize = New-Object System.Drawing.Size(880, 560)
$f.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
$f.TopMost = $true
$f.KeyPreview = $true
$tip = New-Object System.Windows.Forms.ToolTip
$tip.AutoPopDelay = 20000

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.ColumnCount = 2
$root.RowCount = 3
$root.Padding = New-Object System.Windows.Forms.Padding(10)
[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 57)))
[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 43)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

$head = New-Object System.Windows.Forms.Label
$head.Text = $summary
$head.AutoSize = $true
$head.MaximumSize = New-Object System.Drawing.Size(1000, 0)
$head.Margin = New-Object System.Windows.Forms.Padding(2, 0, 2, 8)
$root.Controls.Add($head, 0, 0)
$root.SetColumnSpan($head, 2)

# ---------------- sheets ----------------
$gs = New-Object System.Windows.Forms.GroupBox
$gs.Text = 'Sheets (used from the top down)'
$gs.Dock = 'Fill'
$sl = New-Object System.Windows.Forms.TableLayoutPanel
$sl.Dock = 'Fill'
$sl.ColumnCount = 1
$sl.RowCount = 4
[void]$sl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$sl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$sl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$sl.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$gs.Controls.Add($sl)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.AllowUserToAddRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BackgroundColor = [System.Drawing.SystemColors]::Window
$cUse = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$cUse.HeaderText = 'Use'
$cUse.FillWeight = 12
$cName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$cName.HeaderText = 'Sheet'
$cName.FillWeight = 46
$cW = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$cW.HeaderText = 'Width'
$cW.FillWeight = 16
$cH = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$cH.HeaderText = 'Height'
$cH.FillWeight = 16
$cQ = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$cQ.HeaderText = 'Qty'
$cQ.FillWeight = 14
$cQ.ToolTipText = '0 = as many as needed'
[void]$grid.Columns.AddRange(@($cUse, $cName, $cW, $cH, $cQ))
$sl.Controls.Add($grid, 0, 0)

function AddRow($use, $name, $w, $h, $q, $src) {
  $i = $grid.Rows.Add(@(($use -eq '1'), $name, $w, $h, $q))
  $r = $grid.Rows[$i]
  $r.Tag = $src
  if (([string]$src).StartsWith('D')) {
    $r.Cells[2].ReadOnly = $true
    $r.Cells[3].ReadOnly = $true
    $r.Cells[4].ReadOnly = $true
    $r.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkSlateBlue
  }
  return $i
}
foreach ($s in $sheets) { [void](AddRow $s[0] $s[1] $s[2] $s[3] $s[4] $s[5]) }

function Btn([string]$t, $panel) {
  $b = New-Object System.Windows.Forms.Button
  $b.Text = $t
  $b.AutoSize = $true
  $b.Margin = New-Object System.Windows.Forms.Padding(0, 4, 6, 0)
  $panel.Controls.Add($b)
  return $b
}
function Lbl([string]$t, $panel) {
  $l = New-Object System.Windows.Forms.Label
  $l.Text = $t
  $l.AutoSize = $true
  $l.Margin = New-Object System.Windows.Forms.Padding(0, 9, 4, 0)
  $panel.Controls.Add($l)
  return $l
}

# a common size in one click
$sa = New-Object System.Windows.Forms.FlowLayoutPanel
$sa.Dock = 'Fill'
$sa.AutoSize = $true
$sa.WrapContents = $true
[void](Lbl 'Add a sheet:' $sa)
$cSize = New-Object System.Windows.Forms.ComboBox
$cSize.DropDownStyle = 'DropDownList'
$cSize.Width = 150
$cSize.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
foreach ($z in @('3050 x 2050', '3050 x 1500', '3000 x 1500', '2440 x 1220', '2500 x 1250', '2000 x 1000', '1220 x 2440')) { [void]$cSize.Items.Add($z) }
if ($pageW -ne '' -and $pageH -ne '') {
  $pw = 0.0; $ph = 0.0
  if ([double]::TryParse($pageW, [System.Globalization.NumberStyles]::Float, $inv, [ref]$pw) -and
      [double]::TryParse($pageH, [System.Globalization.NumberStyles]::Float, $inv, [ref]$ph)) {
    [void]$cSize.Items.Add(('Page ' + [Math]::Round($pw).ToString($inv) + ' x ' + [Math]::Round($ph).ToString($inv)))
  }
}
$cSize.SelectedIndex = 0
$sa.Controls.Add($cSize)
$bAdd = Btn 'Add' $sa
$tip.SetToolTip($cSize, 'Pick a common sheet size and press Add. The size can be edited in the list afterwards.')
$sl.Controls.Add($sa, 0, 1)

$sb = New-Object System.Windows.Forms.FlowLayoutPanel
$sb.Dock = 'Fill'
$sb.AutoSize = $true
$sb.WrapContents = $true
$bDel = Btn 'Remove' $sb
$bUp = Btn 'Up' $sb
$bDown = Btn 'Down' $sb
$bAllOn = Btn 'Use all' $sb
$bAllOff = Btn 'Use none' $sb
[void](Lbl 'Qty for all:' $sb)
$tq = New-Object System.Windows.Forms.TextBox
$tq.Width = 50
$tq.Margin = New-Object System.Windows.Forms.Padding(0, 5, 4, 0)
$sb.Controls.Add($tq)
$bQ = Btn 'Apply' $sb
$sl.Controls.Add($sb, 0, 2)

$hint = New-Object System.Windows.Forms.Label
$hint.AutoSize = $true
$hint.MaximumSize = New-Object System.Drawing.Size(560, 0)
$hint.ForeColor = [System.Drawing.Color]::DimGray
$hint.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
$hint.Text = 'Qty 0 = as many as needed. Drawn offcuts (blue) are shapes named "offcut" or "sheet" selected with the parts; the main menu can mark them. Untick a sheet to leave it out.'
$sl.Controls.Add($hint, 0, 3)
$root.Controls.Add($gs, 0, 1)

$bAdd.Add_Click({
  $t = [string]$cSize.SelectedItem
  $m = [regex]::Match($t, '(\d+) x (\d+)')
  if ($m.Success) {
    $w = $m.Groups[1].Value
    $h = $m.Groups[2].Value
    $i = AddRow '1' ('Sheet ' + $w + ' x ' + $h) $w $h '0' 'L'
    $grid.CurrentCell = $grid.Rows[$i].Cells[1]
  }
})
$bDel.Add_Click({
  if ($grid.CurrentRow -ne $null) { $grid.Rows.RemoveAt($grid.CurrentRow.Index) }
})
function MoveRow([int]$d) {
  if ($grid.CurrentRow -eq $null) { return }
  $i = $grid.CurrentRow.Index
  $j = $i + $d
  if ($j -lt 0 -or $j -ge $grid.Rows.Count) { return }
  $r = $grid.Rows[$i]
  $grid.Rows.RemoveAt($i)
  $grid.Rows.Insert($j, $r)
  $grid.CurrentCell = $grid.Rows[$j].Cells[1]
}
$bUp.Add_Click({ MoveRow -1 })
$bDown.Add_Click({ MoveRow 1 })
$bAllOn.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells[0].Value = $true } })
$bAllOff.Add_Click({ foreach ($r in $grid.Rows) { $r.Cells[0].Value = $false } })
$bQ.Add_Click({
  $v = 0
  if ([int]::TryParse($tq.Text.Trim(), [ref]$v) -and $v -ge 0) {
    [void]$grid.EndEdit()
    foreach ($r in $grid.Rows) { if (-not ([string]$r.Tag).StartsWith('D')) { $r.Cells[4].Value = [string]$v } }
  } else { [void]$tq.Focus() }
})

# ---------------- settings: two tabs ----------------
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tls = @()
foreach ($tn in @('Main', 'More')) {
  $tp = New-Object System.Windows.Forms.TabPage
  $tp.Text = $tn
  $tp.Padding = New-Object System.Windows.Forms.Padding(6)
  $tp.AutoScroll = $true
  $tl = New-Object System.Windows.Forms.TableLayoutPanel
  $tl.ColumnCount = 2
  $tl.AutoSize = $true
  $tl.Dock = 'Top'
  [void]$tl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 55)))
  [void]$tl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 45)))
  $tp.Controls.Add($tl)
  $tabs.TabPages.Add($tp)
  $tls += $tl
}
$ctl = @{}
$rows = @(0, 0)
foreach ($fd in $fields) {
  $key = $fd[0]
  $tl = $tls[$fd[4]]
  $row = $rows[$fd[4]]
  $v = ''
  if ($vals.ContainsKey($key)) { $v = $vals[$key] }
  if ($fd[2] -eq 'chk') {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $fd[1]
    $c.AutoSize = $true
    $c.Checked = ($v -eq '1')
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 2)
    $tl.Controls.Add($c, 0, $row)
    $tl.SetColumnSpan($c, 2)
    $tip.SetToolTip($c, $fd[5])
  } else {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $fd[1]
    $l.AutoSize = $true
    $l.Margin = New-Object System.Windows.Forms.Padding(3, 7, 3, 2)
    $tl.Controls.Add($l, 0, $row)
    $tip.SetToolTip($l, $fd[5])
    if ($fd[2] -eq 'combo') {
      $c = New-Object System.Windows.Forms.ComboBox
      $c.DropDownStyle = 'DropDownList'
      $c.Dock = 'Fill'
      $codes = New-Object System.Collections.Generic.List[string]
      foreach ($o in $fd[3].Split('|')) {
        $k = $o.IndexOf('=')
        $codes.Add($o.Substring(0, $k))
        [void]$c.Items.Add($o.Substring($k + 1))
      }
      $c.Tag = $codes
      $sel = 0
      for ($i = 0; $i -lt $codes.Count; $i++) {
        $a = 0.0; $b = 0.0
        if ($codes[$i] -eq $v) { $sel = $i }
        elseif ([double]::TryParse($codes[$i], [System.Globalization.NumberStyles]::Float, $inv, [ref]$a) -and
                [double]::TryParse($v, [System.Globalization.NumberStyles]::Float, $inv, [ref]$b) -and
                [math]::Abs($a - $b) -lt 0.0001) { $sel = $i }
      }
      $c.SelectedIndex = $sel
    } else {
      $c = New-Object System.Windows.Forms.TextBox
      $c.Text = $v
      $c.Dock = 'Fill'
    }
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 2)
    $tl.Controls.Add($c, 1, $row)
    $tip.SetToolTip($c, $fd[5])
  }
  $ctl[$key] = $c
  $rows[$fd[4]] = $row + 1
}
$tls[0].RowCount = $rows[0]
$tls[1].RowCount = $rows[1]
$root.Controls.Add($tabs, 1, 1)

function GetVal([string]$key) {
  $c = $ctl[$key]
  if ($c -is [System.Windows.Forms.CheckBox]) { if ($c.Checked) { return '1' } else { return '0' } }
  if ($c -is [System.Windows.Forms.ComboBox]) { return $c.Tag[$c.SelectedIndex] }
  return $c.Text.Trim().Replace(',', '.')
}
function SetVal([string]$key, [string]$v) {
  $c = $ctl[$key]
  if ($c -eq $null) { return }
  if ($c -is [System.Windows.Forms.CheckBox]) { $c.Checked = ($v -eq '1'); return }
  if ($c -is [System.Windows.Forms.ComboBox]) {
    for ($i = 0; $i -lt $c.Tag.Count; $i++) { if ($c.Tag[$i] -eq $v) { $c.SelectedIndex = $i } }
    return
  }
  $c.Text = $v
}
function Blob() {
  $parts = @()
  foreach ($fd in $fields) { $parts += ($fd[0] + '=' + (GetVal $fd[0])) }
  return ($parts -join ';')
}

# ---------------- bottom bar ----------------
$bar = New-Object System.Windows.Forms.TableLayoutPanel
$bar.Dock = 'Fill'
$bar.AutoSize = $true
$bar.ColumnCount = 2
[void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$left = New-Object System.Windows.Forms.FlowLayoutPanel
$left.AutoSize = $true
$left.Dock = 'Fill'
[void](Lbl 'Preset:' $left)
$cp = New-Object System.Windows.Forms.ComboBox
$cp.Width = 170
$cp.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
foreach ($k in $presets.Keys) { [void]$cp.Items.Add($k) }
$left.Controls.Add($cp)
$bLoad = Btn 'Load' $left
$bSave = Btn 'Save' $left
$bPDel = Btn 'Delete' $left
$tip.SetToolTip($cp, 'Type a name and press Save to keep every setting under it, e.g. Laser 3mm or CNC 6mm.')
$right = New-Object System.Windows.Forms.FlowLayoutPanel
$right.AutoSize = $true
$right.Dock = 'Fill'
$right.FlowDirection = 'RightToLeft'
$bCancel = Btn 'Cancel' $right
$bNest = Btn 'Nest now' $right
$bPreview = Btn 'Preview' $right
$bPreview.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$tip.SetToolTip($bPreview, 'Works the layout out and shows it before anything moves. Apply it there, or come back and change a setting.')
$tip.SetToolTip($bNest, 'Works the layout out and places the parts straight away. Ctrl+Z undoes it.')
$bar.Controls.Add($left, 0, 0)
$bar.Controls.Add($right, 1, 0)
$root.Controls.Add($bar, 0, 2)
$root.SetColumnSpan($bar, 2)
$f.Controls.Add($root)
$f.AcceptButton = $bPreview
$f.CancelButton = $bCancel

$bLoad.Add_Click({
  $n = $cp.Text.Trim()
  if ($presets.Contains($n)) {
    foreach ($kv in ([string]$presets[$n]).Split(';')) {
      $i = $kv.IndexOf('=')
      if ($i -gt 0) { SetVal $kv.Substring(0, $i) $kv.Substring($i + 1) }
    }
  }
})
$bSave.Add_Click({
  $n = $cp.Text.Trim().Replace('|', ' ').Replace(';', ' ').Replace('=', ' ')
  if ($n.Length -eq 0) { [void]$cp.Focus(); return }
  if (-not $presets.Contains($n)) { [void]$cp.Items.Add($n) }
  $presets[$n] = Blob
  $cp.Text = $n
})
$bPDel.Add_Click({
  $n = $cp.Text.Trim()
  if ($presets.Contains($n)) { $presets.Remove($n); $cp.Items.Remove($n); $cp.Text = '' }
})

$script:act = 'CANCEL'
function Check() {
  [void]$grid.EndEdit()
  foreach ($fd in $fields) {
    if ($fd[2] -eq 'num' -or $fd[2] -eq 'int') {
      $t = GetVal $fd[0]
      $x = 0.0
      if (-not [double]::TryParse($t, [System.Globalization.NumberStyles]::Float, $inv, [ref]$x) -or $x -lt 0) {
        $tabs.SelectedIndex = $fd[4]
        [void][System.Windows.Forms.MessageBox]::Show($f, ($fd[1] + ': enter a number of 0 or more.'), $title)
        [void]$ctl[$fd[0]].Focus()
        return $false
      }
    }
  }
  $used = 0
  foreach ($r in $grid.Rows) {
    if ($r.Cells[0].Value -eq $true) {
      $w = 0.0; $h = 0.0; $q = 0
      if (-not [double]::TryParse(([string]$r.Cells[2].Value).Replace(',', '.'), [System.Globalization.NumberStyles]::Float, $inv, [ref]$w) -or $w -le 0 -or
          -not [double]::TryParse(([string]$r.Cells[3].Value).Replace(',', '.'), [System.Globalization.NumberStyles]::Float, $inv, [ref]$h) -or $h -le 0 -or
          -not [int]::TryParse(([string]$r.Cells[4].Value).Trim(), [ref]$q) -or $q -lt 0) {
        [void][System.Windows.Forms.MessageBox]::Show($f, ('Row ' + ($r.Index + 1) + ': width and height must be above 0, qty 0 or more.'), $title)
        $grid.CurrentCell = $r.Cells[2]
        return $false
      }
      $used++
    }
  }
  if ($used -eq 0) {
    [void][System.Windows.Forms.MessageBox]::Show($f, 'Tick at least one sheet.', $title)
    return $false
  }
  return $true
}
$bPreview.Add_Click({ if (Check) { $script:act = 'PREVIEW'; $f.Close() } })
$bNest.Add_Click({ if (Check) { $script:act = 'NEST'; $f.Close() } })
$bCancel.Add_Click({ $script:act = 'CANCEL'; $f.Close() })

[void]$f.ShowDialog()

$o = New-Object System.Collections.Generic.List[string]
$o.Add('A|' + $script:act)
if ($script:act -ne 'CANCEL') {
  foreach ($fd in $fields) { $o.Add('K|' + $fd[0] + '|' + (GetVal $fd[0])) }
  foreach ($r in $grid.Rows) {
    $u = '0'
    if ($r.Cells[0].Value -eq $true) { $u = '1' }
    $nm = ([string]$r.Cells[1].Value).Replace('|', ' ')
    $o.Add('S|' + $u + '|' + $nm + '|' + ([string]$r.Cells[2].Value).Replace(',', '.') + '|' +
           ([string]$r.Cells[3].Value).Replace(',', '.') + '|' + ([string]$r.Cells[4].Value) + '|' + [string]$r.Tag)
  }
}
foreach ($k in $presets.Keys) { $o.Add('R|' + $k + '|' + $presets[$k]) }
[System.IO.File]::WriteAllLines($Out, $o, (New-Object System.Text.UTF8Encoding($true)))
