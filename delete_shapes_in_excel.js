const fs = require('fs')
const path = require('path')
const { spawnSync } = require('child_process')

function parseArgs(argv) {
  const args = {}
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i]
    if (a.startsWith('--')) {
      const eq = a.indexOf('=')
      if (eq > -1) {
        const k = a.slice(2, eq)
        const v = a.slice(eq + 1)
        args[k] = v
      } else {
        const k = a.slice(2)
        const next = argv[i + 1]
        if (next && !next.startsWith('--')) {
          args[k] = next
          i++
        } else {
          args[k] = true
        }
      }
    }
  }
  return args
}

function runExcelCOMDelete({ excelPath, sheetName }) {
  const ps = `
  $ErrorActionPreference = 'Stop'
  $excelPath = '${excelPath.replace(/'/g, "''")}'
  $sheetName = '${(sheetName || '').replace(/'/g, "''")}'
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($excelPath)
  try {
    $targets = @()
    if ($sheetName -and $sheetName.Length -gt 0) {
      $ws = $null
      foreach ($w in $wb.Worksheets) { if ($w.Name -eq $sheetName) { $ws = $w; break } }
      if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
      $targets = @($ws)
    } else {
      $targets = $wb.Worksheets
    }
    $totalBefore = 0; $totalDeleted = 0; $totalRemaining = 0; $fail = 0; $totalSkippedPictures = 0
    foreach ($ws in $targets) {
      try {
        $countBefore = [int]$ws.Shapes.Count
        $deleted = 0; $skippedPictures = 0
        $maxRounds = 5
        for ($round = 1; $round -le $maxRounds; $round++) {
          $ungrouped = 0
          $c = [int]$ws.Shapes.Count
          for ($i = $c; $i -ge 1; $i--) {
            try {
              $s = $ws.Shapes.Item($i)
              $t = [int]$s.Type
              if ($t -eq 6) { try { $s.Ungroup() | Out-Null; $ungrouped++ } catch {} }
            } catch {}
          }
          if ($ungrouped -eq 0) { break }
        }
        $c2 = [int]$ws.Shapes.Count
        for ($i = $c2; $i -ge 1; $i--) {
          try {
            $s = $ws.Shapes.Item($i)
            $t = [int]$s.Type
            if ($t -eq 13 -or $t -eq 11) { $skippedPictures++; continue }
            $s.Delete(); $deleted++
          } catch { $fail++ }
        }
        $countAfter = [int]$ws.Shapes.Count
        $totalBefore += $countBefore
        $totalDeleted += $deleted
        $totalRemaining += $countAfter
        $totalSkippedPictures += $skippedPictures
        Write-Output ("[OK] {0}: deleted={1}, remaining={2}, skippedPictures={3}" -f $ws.Name, $deleted, $countAfter, $skippedPictures)
      } catch { $fail++; Write-Output ("[FAIL] {0}: $($_.Exception.Message)" -f $ws.Name) }
    }
    try { $wb.Save() } catch {}
    Write-Output ("Summary: before={0}, deleted={1}, remaining={2}, skippedPictures={3}, fail={4}" -f $totalBefore, $totalDeleted, $totalRemaining, $totalSkippedPictures, $fail)
  } finally {
    $wb.Close($true)
    $excel.Quit()
  }
  `
  const encoded = Buffer.from(ps, 'utf16le').toString('base64')
  const r = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded], { stdio: 'inherit' })
  if (r.error) throw r.error
  if (r.status !== 0) throw new Error('PowerShell exited with code ' + r.status)
}

function main() {
  const args = parseArgs(process.argv)
  const excelPath = args.excel || 'X\\Users\\jacky\\Downloads\\BRA_006-03_農業経営統計調査（自社内）_テストエビデンス_20251017.xls'
  const sheetName = args.sheet
  const dryRun = !!args.dryRun

  if (dryRun) {
    console.log('干跑模式')
    console.log('EXCEL: ' + excelPath)
    console.log('SHEET: ' + (sheetName || '全部'))
    process.exit(0)
  }
  if (!fs.existsSync(excelPath)) {
    console.error('目标EXCEL不存在: ' + excelPath)
    process.exit(1)
  }
  const backupPath = createBackup(excelPath)
  console.log('已备份: ' + backupPath)
  runExcelCOMDelete({ excelPath, sheetName })
  console.log('完成')
}

main()

function createBackup(src) {
  const dir = path.dirname(src)
  const ext = path.extname(src)
  const base = path.basename(src, ext)
  const d = new Date()
  const stamp = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}_${String(d.getHours()).padStart(2,'0')}${String(d.getMinutes()).padStart(2,'0')}${String(d.getSeconds()).padStart(2,'0')}`
  const dest = path.join(dir, `${base}_backup_${stamp}${ext}`)
  fs.copyFileSync(src, dest)
  return dest
}
