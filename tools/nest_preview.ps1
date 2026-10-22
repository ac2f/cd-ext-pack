param([string]$In, [string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
# written by VBA Print #, so in the code page of this machine
$all = [System.IO.File]::ReadAllLines($In, [System.Text.Encoding]::Default)
$title = 'Nesting - preview'
$report = New-Object System.Collections.Generic.List[string]
$sheets = New-Object System.Collections.Generic.List[object]
$byNo = @{}
$byPart = @{}

# "x y x y ..." in tenths of a millimetre -> PointF[] in millimetres
function Pts([string]$s) {
  $a = $s.Split(' ')
  $n = [int][Math]::Floor($a.Length / 2)
  $r = New-Object 'System.Drawing.PointF[]' $n
  for ($i = 0; $i -lt $n; $i++) {
    $r[$i] = [System.Drawing.PointF]::new(([single]([int]$a[2 * $i]) / 10), ([single]([int]$a[2 * $i + 1]) / 10))
  }
  return ,$r
}

foreach ($ln in $all) {
  $p = $ln.Split('|')
  switch ($p[0]) {
    'T' { $title = $p[1] }
    'R' { $report.Add($p[1]) }
    'S' {
      if ($p.Length -ge 12) {
        $s = New-Object PSObject -Property @{
          No = [int]$p[1]; Name = $p[2]; Parts = [int]$p[3]
          X0 = [int]$p[4] / 10; Y0 = [int]$p[5] / 10; X1 = [int]$p[6] / 10; Y1 = [int]$p[7] / 10
          UX0 = [int]$p[8] / 10; UY0 = [int]$p[9] / 10; UX1 = [int]$p[10] / 10; UY1 = [int]$p[11] / 10
          Outl = (New-Object System.Collections.Generic.List[object])
          Items = (New-Object System.Collections.Generic.List[object])
        }
        $sheets.Add($s)
        $byNo[$s.No] = $s
      }
    }
    'O' {
      if ($p.Length -ge 4 -and $byNo.ContainsKey([int]$p[1])) {
        $pts = Pts $p[3]
        if ($pts.Length -ge 3) { $byNo[[int]$p[1]].Outl.Add($pts) }
      }
    }
    'P' {
      if ($p.Length -ge 6 -and $byNo.ContainsKey([int]$p[1])) {
        $pk = [int]$p[2]
        $it = $byPart[$pk]
        if ($it -eq $null) {
          $it = New-Object PSObject -Property @{
            Cls = [int]$p[3]
            Closed = (New-Object System.Collections.Generic.List[object])
            Open = (New-Object System.Collections.Generic.List[object])
          }
          $byPart[$pk] = $it
          $byNo[[int]$p[1]].Items.Add($it)
        }
        $pts = Pts $p[5]
        if ($p[4] -eq '1') {
          if ($pts.Length -ge 2) { $it.Open.Add($pts) }
        } elseif ($pts.Length -ge 3) {
          $it.Closed.Add($pts)
        }
      }
    }
  }
}

$f = New-Object System.Windows.Forms.Form
$f.Text = $title
$f.StartPosition = 'CenterScreen'
$f.Size = New-Object System.Drawing.Size(1180, 760)
$f.MinimumSize = New-Object System.Drawing.Size(820, 520)
$f.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
$f.TopMost = $true
$f.KeyPreview = $true

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.ColumnCount = 2
$root.RowCount = 2
$root.Padding = New-Object System.Windows.Forms.Padding(10)
[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 68)))
[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 32)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

$pb = New-Object System.Windows.Forms.PictureBox
$pb.Dock = 'Fill'
$pb.BorderStyle = 'FixedSingle'
$pb.BackColor = [System.Drawing.Color]::White
$root.Controls.Add($pb, 0, 0)

$side = New-Object System.Windows.Forms.TableLayoutPanel
$side.Dock = 'Fill'
$side.ColumnCount = 1
$side.RowCount = 4
[void]$side.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$side.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 35)))
[void]$side.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$side.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 65)))
$l1 = New-Object System.Windows.Forms.Label
$l1.Text = 'Show'
$l1.AutoSize = $true
$l1.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 3)
$side.Controls.Add($l1, 0, 0)
$lst = New-Object System.Windows.Forms.ListBox
$lst.Dock = 'Fill'
$lst.IntegralHeight = $false
[void]$lst.Items.Add('All sheets')
foreach ($s in $sheets) { [void]$lst.Items.Add(('Sheet ' + $s.No + ': ' + $s.Name + ' (' + $s.Parts + ' parts)')) }
$side.Controls.Add($lst, 0, 1)
$l2 = New-Object System.Windows.Forms.Label
$l2.Text = 'Report'
$l2.AutoSize = $true
$l2.Margin = New-Object System.Windows.Forms.Padding(3, 8, 3, 3)
$side.Controls.Add($l2, 0, 2)
$txt = New-Object System.Windows.Forms.TextBox
$txt.Multiline = $true
$txt.ReadOnly = $true
$txt.ScrollBars = 'Both'
$txt.WordWrap = $false
$txt.Dock = 'Fill'
$txt.Font = New-Object System.Drawing.Font('Consolas', 9)
$txt.Text = [string]::Join([Environment]::NewLine, $report.ToArray())
$side.Controls.Add($txt, 0, 3)
$root.Controls.Add($side, 1, 0)

$bar = New-Object System.Windows.Forms.TableLayoutPanel
$bar.Dock = 'Fill'
$bar.AutoSize = $true
$bar.ColumnCount = 2
[void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$note = New-Object System.Windows.Forms.Label
$note.AutoSize = $true
$note.ForeColor = [System.Drawing.Color]::DimGray
$note.Margin = New-Object System.Windows.Forms.Padding(2, 10, 2, 0)
$note.Text = 'Nothing has moved yet. Apply places the parts as shown (Ctrl+Z undoes it).'
$bar.Controls.Add($note, 0, 0)
$right = New-Object System.Windows.Forms.FlowLayoutPanel
$right.AutoSize = $true
$right.Dock = 'Fill'
$right.FlowDirection = 'RightToLeft'
function Btn([string]$t) {
  $b = New-Object System.Windows.Forms.Button
  $b.Text = $t
  $b.AutoSize = $true
  $b.Margin = New-Object System.Windows.Forms.Padding(0, 4, 6, 0)
  $right.Controls.Add($b)
  return $b
}
$bCancel = Btn 'Cancel'
$bBack = Btn 'Change settings'
$bApply = Btn 'Apply'
$bApply.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
$bar.Controls.Add($right, 1, 0)
$root.Controls.Add($bar, 0, 1)
$root.SetColumnSpan($bar, 2)
$f.Controls.Add($root)
$f.AcceptButton = $bApply
$f.CancelButton = $bCancel

# ---------------- drawing ----------------
$colors = @(
  @(166, 206, 227), @(178, 223, 138), @(251, 180, 174), @(253, 205, 172),
  @(203, 213, 232), @(244, 202, 228), @(230, 245, 201), @(255, 242, 174),
  @(204, 235, 197), @(222, 203, 228)
)
$brushes = @()
foreach ($c in $colors) { $brushes += (New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, $c[0], $c[1], $c[2]))) }
$brSheet = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 236, 238, 241))
$brUse = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 250, 250, 250))
$brText = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 40, 40, 40))
$penSheet = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 120, 120, 120), 1)
$penUse = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 150, 180, 220), 1)
$penPart = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 45, 60, 85), 1)
$fLabel = New-Object System.Drawing.Font('Segoe UI', 9)

function Render {
  try {
    $w = $pb.ClientSize.Width
    $h = $pb.ClientSize.Height
    if ($w -lt 40 -or $h -lt 40) { return }
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::White)
    $show = New-Object System.Collections.Generic.List[object]
    if ($lst.SelectedIndex -le 0) {
      foreach ($s in $sheets) { $show.Add($s) }
    } else {
      $show.Add($sheets[$lst.SelectedIndex - 1])
    }
    if ($show.Count -eq 0) {
      $g.DrawString('Nothing was placed.', $fLabel, $brText, [single]12, [single]12)
    } else {
      $minx = $show[0].X0; $miny = $show[0].Y0; $maxx = $show[0].X1; $maxy = $show[0].Y1
      foreach ($s in $show) {
        if ($s.X0 -lt $minx) { $minx = $s.X0 }
        if ($s.Y0 -lt $miny) { $miny = $s.Y0 }
        if ($s.X1 -gt $maxx) { $maxx = $s.X1 }
        if ($s.Y1 -gt $maxy) { $maxy = $s.Y1 }
      }
      $bw = [Math]::Max(1.0, $maxx - $minx)
      $bh = [Math]::Max(1.0, $maxy - $miny)
      $sc = [Math]::Min(($w - 30) / $bw, ($h - 50) / $bh)
      $ox = ($w - $bw * $sc) / 2 - $minx * $sc
      $oy = ($h + $bh * $sc) / 2 + 10 + $miny * $sc
      # page y points up, the screen down
      $g.TranslateTransform([single]$ox, [single]$oy)
      $g.ScaleTransform([single]$sc, [single](-$sc))
      $thin = [single](1.0 / $sc)
      $penSheet.Width = $thin
      $penUse.Width = $thin
      $penPart.Width = $thin
      foreach ($s in $show) {
        if ($s.Outl.Count -gt 0) {
          $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
          $gp.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate
          foreach ($o in $s.Outl) { $gp.AddPolygon($o) }
          $g.FillPath($brUse, $gp)
          $g.DrawPath($penSheet, $gp)
          $gp.Dispose()
        } else {
          $g.FillRectangle($brSheet, [single]$s.X0, [single]$s.Y0, [single]($s.X1 - $s.X0), [single]($s.Y1 - $s.Y0))
          $g.DrawRectangle($penSheet, [single]$s.X0, [single]$s.Y0, [single]($s.X1 - $s.X0), [single]($s.Y1 - $s.Y0))
          if ($s.UX1 -gt $s.UX0 -and $s.UY1 -gt $s.UY0) {
            $g.FillRectangle($brUse, [single]$s.UX0, [single]$s.UY0, [single]($s.UX1 - $s.UX0), [single]($s.UY1 - $s.UY0))
            $g.DrawRectangle($penUse, [single]$s.UX0, [single]$s.UY0, [single]($s.UX1 - $s.UX0), [single]($s.UY1 - $s.UY0))
          }
        }
        foreach ($it in $s.Items) {
          if ($it.Closed.Count -gt 0) {
            $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
            $gp.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate
            foreach ($c in $it.Closed) { $gp.AddPolygon($c) }
            $g.FillPath($brushes[$it.Cls % $brushes.Count], $gp)
            $g.DrawPath($penPart, $gp)
            $gp.Dispose()
          }
          foreach ($c in $it.Open) { $g.DrawLines($penPart, $c) }
        }
      }
      $g.ResetTransform()
      foreach ($s in $show) {
        $tx = $ox + $s.X0 * $sc
        $ty = $oy - $s.Y1 * $sc - 18
        $g.DrawString(([string]$s.No + '  ' + $s.Name), $fLabel, $brText, [single]$tx, [single]$ty)
      }
    }
    $g.Dispose()
    $old = $pb.Image
    $pb.Image = $bmp
    if ($old -ne $null) { $old.Dispose() }
  } catch {
    # a drawing problem must not cost the user the window
  }
}

$lst.Add_SelectedIndexChanged({ Render })
$pb.Add_Resize({ Render })
$f.Add_Shown({ Render })
$lst.SelectedIndex = 0

$script:act = 'CANCEL'
$bApply.Add_Click({ $script:act = 'APPLY'; $f.Close() })
$bBack.Add_Click({ $script:act = 'SETTINGS'; $f.Close() })
$bCancel.Add_Click({ $script:act = 'CANCEL'; $f.Close() })

[void]$f.ShowDialog()

$o = New-Object System.Collections.Generic.List[string]
$o.Add('A|' + $script:act)
[System.IO.File]::WriteAllLines($Out, $o, (New-Object System.Text.UTF8Encoding($true)))
