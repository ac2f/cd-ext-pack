#!/usr/bin/env python3
"""ac2f pack - nesting pencerelerini (PowerShell + WinForms) gercekten calistirir.

CorelDRAW'in actigi iki pencere, tools/nest_window.ps1 (ayarlar) ve
tools/nest_preview.ps1 (onizleme), Linux'taki PowerShell 7 ile ve
tools/winstub.cs icindeki WinForms/GDI+ taklidiyle calistirilir. Sinama
dugmelere basar, kutulara yazar, sonra pencerenin yazdigi cikis dosyasina
bakar. Ayrica:

  - betik ASCII mi (VBA'ya Print # ile gomulur)
  - Windows PowerShell 5.1'in tanimadigi bir yapi var mi (?:, ??, ?., &&)
  - salt-okunur otomatik degisken ($PID, $Input ...) atanmis mi

PowerShell yolu: AC2F_PWSH ortam degiskeni ya da PATH'teki pwsh. Bulunamazsa
sinama atlanir (cikis kodu 0, uyari basilir).

Kullanim:  python3 tools/test_windows.py
"""
import os, shutil, subprocess, sys, tempfile, time

HERE = os.path.dirname(os.path.abspath(__file__))
STUB = os.path.join(HERE, 'winstub.cs')
SETTINGS = os.path.join(HERE, 'nest_window.ps1')
PREVIEW = os.path.join(HERE, 'nest_preview.ps1')
FAILS = []

DRIVER = r'''
param([string]$Script, [string]$In, [string]$Out, [string]$Stub, [string]$Actions)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition ([System.IO.File]::ReadAllText($Stub)) -ReferencedAssemblies @(
  'System.Drawing.Primitives', 'System.Collections', 'System.Runtime', 'System.Linq')
function Add-Type {
  param([string]$AssemblyName)
  if (@('System.Windows.Forms', 'System.Drawing') -contains $AssemblyName) { return }
  throw ('unexpected Add-Type ' + $AssemblyName)
}
$src = [System.IO.File]::ReadAllText($Script)
$src = [regex]::Replace($src, 'System\.Drawing\.(Font|FontStyle|Bitmap|Graphics|Pen|SolidBrush|SystemColors|Brushes|Image|Drawing2D)\b', 'StubDrawing.$1')
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ac2f_stubbed_' + [Guid]::NewGuid().ToString('N') + '.ps1')
[System.IO.File]::WriteAllText($tmp, $src)
function Ctl($f, [string]$text) {
  $c = $f.Find($text)
  if ($c -eq $null) { throw ('no control with the text ' + $text) }
  return $c
}
function Click($f, [string]$text) { (Ctl $f $text).PerformClick() }
$OnShow = { param($f) }
. $Actions
[System.Windows.Forms.Harness]::OnShow = [Action[System.Windows.Forms.Form]]$OnShow
[StubDrawing.Stats]::Reset()
try {
  & $tmp -In $In -Out $Out
} finally {
  Remove-Item $tmp -ErrorAction SilentlyContinue
}
'STATS|' + [StubDrawing.Stats]::Renders + '|' + [StubDrawing.Stats]::Polygons + '|' + [StubDrawing.Stats]::Lines + '|' + [StubDrawing.Stats]::Strings
foreach ($m in [System.Windows.Forms.Harness]::Messages) { 'MSG|' + $m }
foreach ($m in [System.Windows.Forms.Harness]::Notes) { 'NOTE|' + $m }
foreach ($t in [StubDrawing.Stats]::Text) { 'TEXT|' + $t }
'''

COMPAT = r'''
param([string]$Path)
$tok = $null; $err = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tok, [ref]$err)
foreach ($e in $err) { 'PARSE|' + $e.Extent.StartLineNumber + '|' + $e.Message }
$bad = $ast.FindAll({ param($n)
  $n -is [System.Management.Automation.Language.TernaryExpressionAst] -or
  $n -is [System.Management.Automation.Language.PipelineChainAst] -or
  ($n -is [System.Management.Automation.Language.BinaryExpressionAst] -and $n.Operator -eq 'QuestionQuestion') -or
  ($n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Operator -eq 'QuestionQuestionEquals') -or
  ($n -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.NullConditional)
}, $true)
foreach ($b in $bad) { 'PS7|' + $b.Extent.StartLineNumber + '|' + $b.Extent.Text }
$auto = @('pid', 'input', 'args', 'this', 'host', 'home', 'error', 'matches', 'switch', 'foreach', 'psitem', 'true', 'false', 'null', 'shellid', 'pshome', 'executioncontext', 'myinvocation', 'sender', 'event', 'eventargs', 'profile', 'pwd', 'lastexitcode', 'stacktrace')
$asg = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)
foreach ($a in $asg) {
  $l = $a.Left
  if ($l -is [System.Management.Automation.Language.ConvertExpressionAst]) { $l = $l.Child }
  if ($l -is [System.Management.Automation.Language.VariableExpressionAst]) {
    $n = $l.VariablePath.UserPath.ToLower()
    if ($auto -contains $n) { 'AUTO|' + $a.Extent.StartLineNumber + '|' + $n }
  }
}
$loops = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ForEachStatementAst] }, $true)
foreach ($fe in $loops) {
  $n = $fe.Variable.VariablePath.UserPath.ToLower()
  if ($auto -contains $n) { 'AUTO|' + $fe.Extent.StartLineNumber + '|' + $n }
}
'''


def pwsh_path():
    p = os.environ.get('AC2F_PWSH')
    if p and os.path.exists(p):
        return p
    return shutil.which('pwsh')


def expect(cond, msg):
    print(('  ok   ' if cond else '  FAIL ') + msg)
    if not cond:
        FAILS.append(msg)


def scenario(name):
    print('\n' + name)


class Win(object):
    """Runs one window script with the stub and a block of test actions."""

    def __init__(self, pwsh):
        self.pwsh = pwsh
        self.tmp = tempfile.mkdtemp(prefix='ac2f_win_')
        self.driver = os.path.join(self.tmp, 'driver.ps1')
        with open(self.driver, 'w') as fh:
            fh.write(DRIVER)

    def run(self, script, inp, actions, encoding='utf-8-sig'):
        inf = os.path.join(self.tmp, 'in.txt')
        outf = os.path.join(self.tmp, 'out.txt')
        act = os.path.join(self.tmp, 'actions.ps1')
        with open(inf, 'w', encoding=encoding, newline='') as fh:
            fh.write('\r\n'.join(inp) + '\r\n')
        with open(act, 'w') as fh:
            fh.write(actions)
        if os.path.exists(outf):
            os.remove(outf)
        return self.run_files(script, inf, outf, act)

    def run_files(self, script, inf, outf, act):
        t0 = time.time()
        r = subprocess.run([self.pwsh, '-NoProfile', '-NonInteractive', '-File', self.driver,
                            '-Script', script, '-In', inf, '-Out', outf, '-Stub', STUB, '-Actions', act],
                           capture_output=True, text=True, timeout=300, env=dict(os.environ, NO_COLOR='1'))
        res = {'rc': r.returncode, 'stderr': r.stderr, 'out': None, 'msgs': [], 'notes': [],
               'text': [], 'stats': None, 'secs': time.time() - t0}
        for ln in r.stdout.splitlines():
            if ln.startswith('STATS|'):
                res['stats'] = [int(x) for x in ln.split('|')[1:]]
            elif ln.startswith('MSG|'):
                res['msgs'].append(ln[4:])
            elif ln.startswith('NOTE|'):
                res['notes'].append(ln[5:])
            elif ln.startswith('TEXT|'):
                res['text'].append(ln[5:])
        if os.path.exists(outf):
            res['out'] = open(outf, encoding='utf-8-sig').read().splitlines()
        if r.returncode != 0:
            print(r.stdout[-2000:])
            print(r.stderr[-3000:])
        return res

    def close(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


def compat(pwsh, path):
    tmp = tempfile.mkdtemp(prefix='ac2f_compat_')
    try:
        drv = os.path.join(tmp, 'compat.ps1')
        with open(drv, 'w') as fh:
            fh.write(COMPAT)
        r = subprocess.run([pwsh, '-NoProfile', '-NonInteractive', '-File', drv, '-Path', path],
                           capture_output=True, text=True, timeout=120)
        return [ln for ln in r.stdout.splitlines() if ln.strip()] + ([r.stderr] if r.returncode else [])
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


SETTINGS_IN = [
    'T|ac2f pack - Nesting',
    'P|3 part(s) selected, 1 drawn offcut sheet(s). Preview shows the layout before anything moves.',
    'G|3050|2050',
    'K|gap|4', 'K|mL|10', 'K|mR|10', 'K|mT|10', 'K|mB|10', 'K|om|5', 'K|rot|90', 'K|mirror|0',
    'K|copies|1', 'K|inner|1', 'K|holes|0', 'K|corner|1', 'K|dir|0', 'K|strat|0', 'K|res|1',
    'K|time|10', 'K|group|1', 'K|outline|1', 'K|start|0', 'K|sgap|20',
    'S|1|Big|3050|2050|0|L',
    'S|0|Small|2440|1220|2|L',
    'S|1|Offcut 1: offcut A (850 x 600)|850|600|1|D2',
    'R|Laser 3mm|gap=3;mL=5;mR=5;mT=5;mB=5;rot=45',
]


def kv(out):
    d = {}
    for ln in out or []:
        f = ln.split('|')
        if f[0] == 'K':
            d[f[1]] = f[2]
    return d


def rows(out):
    return [ln.split('|')[1:] for ln in out or [] if ln.startswith('S|')]


def t_settings(w):
    scenario('Ayar penceresi: hic dokunmadan Preview')
    r = w.run(SETTINGS, SETTINGS_IN, r'''
$OnShow = { param($f)
  if ($f.Text -ne 'ac2f pack - Nesting') { throw ('title ' + $f.Text) }
  Click $f 'Preview'
}''')
    expect(r['rc'] == 0, 'betik hatasiz calisti (%.1f s)' % r['secs'])
    out = r['out'] or []
    expect(out[:1] == ['A|PREVIEW'], 'eylem PREVIEW: %s' % out[:1])
    d = kv(out)
    expect(d.get('gap') == '4' and d.get('rot') == '90' and d.get('res') == '1' and d.get('time') == '10',
           'ayarlar aynen geri geldi')
    expect(len(d) == 20, '20 ayar: %d' % len(d))
    rs = rows(out)
    expect(rs == [['1', 'Big', '3050', '2050', '0', 'L'], ['0', 'Small', '2440', '1220', '2', 'L'],
                  ['1', 'Offcut 1: offcut A (850 x 600)', '850', '600', '1', 'D2']], 'tabakalar sirasiyla: %s' % rs)
    expect('R|Laser 3mm|gap=3;mL=5;mR=5;mT=5;mB=5;rot=45' in out, 'hazir ayar geri yazildi')

    scenario('Ayar penceresi: degistir, tabaka ekle, tasi, hazir ayar kaydet, Nest now')
    r = w.run(SETTINGS, SETTINGS_IN, r'''
$OnShow = { param($f)
  $tabs = $f.FirstOf('TabControl')
  if ($tabs.TabPages.Count -ne 2) { throw 'two tabs expected' }
  $main = $tabs.TabPages[0]
  $boxes = $main.AllOf('TextBox')
  $boxes[0].Text = '6,5'
  $combos = $main.AllOf('ComboBox')
  $combos[0].SelectedIndex = 3
  $more = $tabs.TabPages[1]
  (Ctl $more 'Parts may be mirrored').PerformClick()
  $mc = $more.AllOf('ComboBox')
  $mc[0].SelectedIndex = 3
  $size = $null
  foreach ($c in $f.AllOf('ComboBox')) { if ($c.Items.Count -gt 0 -and [string]$c.Items[0] -eq '3050 x 2050') { $size = $c } }
  if ($size -eq $null) { throw 'no size list' }
  if ([string]$size.Items[$size.Items.Count - 1] -ne 'Page 3050 x 2050') { throw ('page item ' + $size.Items[$size.Items.Count - 1]) }
  $size.SelectedIndex = 3
  Click $f 'Add'
  Click $f 'Up'
  Click $f 'Up'
  $q = $null
  foreach ($t in $f.AllOf('TextBox')) { if ($t.Width -eq 50) { $q = $t } }
  $q.Text = '2'
  Click $f 'Apply'
  $preset = $null
  foreach ($c in $f.AllOf('ComboBox')) { if ($c.Width -eq 170) { $preset = $c } }
  $preset.Text = 'CNC 6mm'
  Click $f 'Save'
  Click $f 'Nest now'
}''')
    expect(r['rc'] == 0, 'betik hatasiz calisti')
    out = r['out'] or []
    expect(out[:1] == ['A|NEST'], 'eylem NEST')
    d = kv(out)
    expect(d.get('gap') == '6.5', 'virgullu aralik noktaya cevrildi: %s' % d.get('gap'))
    expect(d.get('rot') == '45', 'donme 45: %s' % d.get('rot'))
    expect(d.get('mirror') == '1', 'ayna acildi')
    expect(d.get('corner') == '4', 'baslangic kosesi sag ust: %s' % d.get('corner'))
    rs = rows(out)
    expect(len(rs) == 4 and rs[1][1:4] == ['Sheet 2440 x 1220', '2440', '1220'],
           'eklenen tabaka iki yukari tasindi: %s' % [x[1] for x in rs])
    expect([x[4] for x in rs] == ['2', '2', '2', '1'], 'herkese adet 2, artik 1 kaldi: %s' % [x[4] for x in rs])
    expect(len(rs) == 4 and rs[3][5] == 'D2', 'artigin kaynagi korundu')
    blob = [ln for ln in out if ln.startswith('R|CNC 6mm|')]
    expect(blob and 'gap=6.5' in blob[0] and 'rot=45' in blob[0] and 'mirror=1' in blob[0],
           'hazir ayar o anki ayarlarla kaydedildi')

    scenario('Ayar penceresi: hatali sayi, tabaka yok, hazir ayar yukle, Cancel')
    r = w.run(SETTINGS, SETTINGS_IN, r'''
$OnShow = { param($f)
  $main = $f.FirstOf('TabControl').TabPages[0]
  $main.AllOf('TextBox')[0].Text = 'abc'
  Click $f 'Preview'
  if ($f.IsClosed) { throw 'closed despite a bad number' }
  $main.AllOf('TextBox')[0].Text = '3'
  Click $f 'Use none'
  Click $f 'Nest now'
  if ($f.IsClosed) { throw 'closed with no sheet ticked' }
  $preset = $null
  foreach ($c in $f.AllOf('ComboBox')) { if ($c.Width -eq 170) { $preset = $c } }
  $preset.Text = 'Laser 3mm'
  Click $f 'Load'
  if ($main.AllOf('TextBox')[0].Text -ne '3') { throw 'preset gap not loaded' }
  Click $f 'Delete'
  Click $f 'Cancel'
}''')
    expect(r['rc'] == 0, 'betik hatasiz calisti')
    out = r['out'] or []
    expect(out[:1] == ['A|CANCEL'], 'eylem CANCEL')
    expect(len(r['msgs']) == 2 and 'number' in r['msgs'][0] and 'Tick' in r['msgs'][1],
           'iki uyari: %s' % r['msgs'])
    expect(not [ln for ln in out if ln.startswith('K|')], 'Cancel ayar yazmaz')
    expect(not [ln for ln in out if ln.startswith('R|Laser')], 'silinen hazir ayar geri yazilmadi')

    scenario('Ayar penceresi: X ile kapatma')
    r = w.run(SETTINGS, SETTINGS_IN, '$OnShow = { param($f) }')
    expect(r['rc'] == 0 and (r['out'] or [])[:1] == ['A|CANCEL'], 'X = CANCEL')


def preview_input():
    lines = ['T|ac2f pack - Nesting - preview',
             'R|PREVIEW - nothing has moved yet', 'R|   Parts             : 3 of 4 placed',
             'S|1|Big|3|0|0|30500|20500|100|100|30400|20400',
             'S|2|Offcut 1: offcut A|1|40000|0|48500|6000|40000|0|48500|6000',
             'O|2|0|40000 0 48500 0 48500 6000 40000 6000',
             'O|2|0|42000 2000 43000 2000 43000 3000',
             'P|1|0|0|0|200 200 2200 200 2200 1500 200 1500',
             'P|1|0|0|0|600 600 900 600 900 900',
             'P|1|1|1|1|3000 3000 5000 3000',
             'P|1|2|1|0|5000 5000 6000 5000 6000 6000',
             'P|2|3|0|0|41000 1000 41500 1000 41500 1500 41000 1500']
    return lines


def t_preview(w):
    scenario('Onizleme: cizim, tabaka secimi, Apply')
    r = w.run(PREVIEW, preview_input(), r'''
$OnShow = { param($f)
  $pb = $f.FirstOf('PictureBox')
  if ($pb.Image -eq $null) { throw 'nothing drawn when shown' }
  $lst = $f.FirstOf('ListBox')
  if ($lst.Items.Count -ne 3) { throw ('list ' + $lst.Items.Count) }
  $first = $pb.Image
  $lst.SelectedIndex = 2
  if ($pb.Image -eq $first) { throw 'not redrawn for one sheet' }
  if (-not $first.Disposed) { throw 'old bitmap kept' }
  $lst.SelectedIndex = 0
  $pb.Size = New-Object System.Drawing.Size(30, 20)
  $pb.Size = New-Object System.Drawing.Size(700, 500)
  $txt = $f.FirstOf('TextBox')
  if (-not $txt.Text.StartsWith('PREVIEW')) { throw 'report text' }
  Click $f 'Apply'
}''', encoding='latin-1')
    expect(r['rc'] == 0, 'betik hatasiz calisti (%.1f s)' % r['secs'])
    expect((r['out'] or [])[:1] == ['A|APPLY'], 'eylem APPLY')
    st = r['stats'] or [0, 0, 0, 0]
    expect(st[0] >= 3, 'en az uc cizim: %s' % st[0])
    expect(st[1] > 0 and st[2] > 0, 'kapali konturlar dolduruldu, acik cizgi cizildi: %s' % st)
    expect(any(t.startswith('1  Big') for t in r['text']) and any(t.startswith('2  Offcut') for t in r['text']),
           'tabaka etiketleri: %s' % sorted(set(r['text']))[:4])

    for button, act in (('Change settings', 'SETTINGS'), ('Cancel', 'CANCEL')):
        scenario('Onizleme: ' + button)
        r = w.run(PREVIEW, preview_input(), "$OnShow = { param($f) Click $f '%s' }" % button, encoding='latin-1')
        expect(r['rc'] == 0 and (r['out'] or [])[:1] == ['A|' + act], 'eylem ' + act)

    scenario('Onizleme: yerlesen yok, X ile kapatma')
    r = w.run(PREVIEW, ['T|x', 'R|nothing'], '$OnShow = { param($f) }', encoding='latin-1')
    expect(r['rc'] == 0 and (r['out'] or [])[:1] == ['A|CANCEL'], 'X = CANCEL')
    expect('Nothing was placed.' in r['text'], 'bos onizleme metni')


def t_static(pwsh):
    scenario('Betikler: ASCII, PowerShell 5.1 uyumu, otomatik degiskenler')
    for p in (SETTINGS, PREVIEW):
        raw = open(p, 'rb').read()
        expect(all(b < 127 for b in raw), '%s ASCII' % os.path.basename(p))
        probs = compat(pwsh, p)
        expect(not probs, '%s: %s' % (os.path.basename(p), probs or 'sorun yok'))


def main():
    pwsh = pwsh_path()
    if not pwsh:
        print('PowerShell (pwsh) bulunamadi: AC2F_PWSH ile yolunu verin. Pencere sinamalari atlandi.')
        return 0
    t0 = time.time()
    t_static(pwsh)
    w = Win(pwsh)
    try:
        t_settings(w)
        t_preview(w)
    finally:
        w.close()
    print('\n%s (%.1f s)' % ('HATA: %d' % len(FAILS) if FAILS else 'Hepsi gecti', time.time() - t0))
    return 1 if FAILS else 0


if __name__ == '__main__':
    sys.exit(main())
