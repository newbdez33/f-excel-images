const fs = require('fs')
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

function runExcelCOMExtract({ excelPath, sheetName, outDir }) {
  const ps = `
  $ErrorActionPreference = 'Stop'
  $excelPath = '${excelPath.replace(/'/g, "''")}'
  $sheetName = '${(sheetName || '').replace(/'/g, "''")}'
  $outDir = '${outDir.replace(/'/g, "''")}'
  New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($excelPath)
  try { Add-Type -AssemblyName System.Windows.Forms } catch {}
  try { Add-Type -AssemblyName System.Drawing } catch {}

  function ExportShapeImage($ws, $shape, $file) {
    $ok = $false
    try {
      $shape.Export($file) | Out-Null
      $img = [System.Drawing.Image]::FromFile($file)
      $w = [int]$img.Width; $h = [int]$img.Height
      $img.Dispose()
      if ($w -gt 0 -and $h -gt 0) { return $true }
    } catch { $ok = $false; Write-Output ("[INFO] Export method failed: {0}:{1} -> {2}" -f $ws.Name, $shape.Name, $_.Exception.Message) }
    try { Remove-Item -Path $file -ErrorAction SilentlyContinue } catch {}
    foreach ($app in @(-4154, 1)) { # xlPrinter, xlScreen
      foreach ($fmt in @(-4147, 2)) { # xlPicture, xlBitmap
        try {
          try { $shape.Select() } catch {}
          $shape.CopyPicture($app, $fmt)
          Start-Sleep -Milliseconds 50
          $before = $ws.Shapes.Count
          try { $ws.Activate() } catch { Write-Output ("[INFO] Activate failed: {0} -> {1}" -f $ws.Name, $_.Exception.Message) }
          $ws.Paste() | Out-Null
          $after = $ws.Shapes.Count
          if ($after -gt $before) {
            $tmp = $ws.Shapes.Item($after)
            try {
              $tmp.Export($file) | Out-Null
              $img = [System.Drawing.Image]::FromFile($file)
              $w = [int]$img.Width; $h = [int]$img.Height
              $img.Dispose()
              if ($w -gt 0 -and $h -gt 0) { try { $tmp.Delete() } catch {}; try { $excel.CutCopyMode = $false } catch {}; return $true }
            } catch { Write-Output ("[INFO] Temp export failed: {0}:{1} -> {2}" -f $ws.Name, $tmp.Name, $_.Exception.Message) }
            try { $tmp.Delete() } catch {}
          }
          try { $excel.CutCopyMode = $false } catch {}
        } catch { $ok = $false; Write-Output ("[INFO] CopyPicture/Paste failed: {0}:{1} -> {2}" -f $ws.Name, $shape.Name, $_.Exception.Message) }
      }
    }
    foreach ($app in @(-4154, 1)) {
      foreach ($fmt in @(-4147, 2)) {
        try {
          try { $shape.Select() } catch {}
          $shape.CopyPicture($app, $fmt)
          Start-Sleep -Milliseconds 50
          $ch = $ws.ChartObjects().Add(0, 0, [double]$shape.Width, [double]$shape.Height)
          $chart = $ch.Chart
          try { $chart.ChartArea.Format.Line.Visible = $false } catch {}
          try { $chart.ChartArea.Interior.Color = 16777215 } catch {}
          try { $chart.PlotArea.Format.Fill.Visible = $false } catch {}
          $chart.Paste() | Out-Null
          $chart.Export($file) | Out-Null
          $img = [System.Drawing.Image]::FromFile($file)
          $w = [int]$img.Width; $h = [int]$img.Height
          $img.Dispose()
          try { $ch.Delete() } catch {}
          try { $excel.CutCopyMode = $false } catch {}
          if ($w -gt 0 -and $h -gt 0) { return $true }
        } catch { Write-Output ("[INFO] Chart export failed: {0}:{1} -> {2}" -f $ws.Name, $shape.Name, $_.Exception.Message) }
      }
    }
    return $false
  }
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
    $ok = 0; $fail = 0
    foreach ($ws in $targets) {
      try {
        if ($ws.ProtectContents -or $ws.ProtectDrawingObjects) {
          try { $ws.Unprotect() | Out-Null; Write-Output ("[INFO] Unprotected sheet: {0}" -f $ws.Name) } catch { Write-Output ("[INFO] Unprotect failed: {0} -> {1}" -f $ws.Name, $_.Exception.Message) }
        }
        $sheetSafe = ($ws.Name -replace '[\\/:*?"<>|]','_')
        $sheetOut = Join-Path $outDir $sheetSafe
        New-Item -ItemType Directory -Path $sheetOut -Force | Out-Null
        Write-Output ("[SHEET] Processing: {0}" -f $ws.Name)
        try { $ws.Activate() } catch {}
        try { $excel.ActiveWindow.ScrollRow = 1 } catch {}
        try { $excel.ActiveWindow.ScrollColumn = 1 } catch {}
        try { $ws.Cells.Item(1,1).Select() } catch {}
        $okSheet = 0; $failSheet = 0
        $sheetIdx = 0
        $shapeRecords = @()
        foreach ($s in $ws.Shapes) {
          try {
            $t = [int]$s.Type
            if ($t -eq 6) {
              foreach ($gi in $s.GroupItems) {
                try {
                  $shapeRecords += [PSCustomObject]@{ S = $gi; Name = $gi.Name; Top = [double]$gi.Top; Left = [double]$gi.Left }
                } catch {}
              }
            } else {
              $shapeRecords += [PSCustomObject]@{ S = $s; Name = $s.Name; Top = [double]$s.Top; Left = [double]$s.Left }
            }
          } catch {}
        }
        $ordered = $shapeRecords | Sort-Object -Property Top, Left, Name
        foreach ($rec in $ordered) {
          try {
            $sheetIdx++
            $file = Join-Path $sheetOut ("$sheetIdx.png")
            $done = ExportShapeImage $ws $rec.S $file
            if ($done) {
              $exists = Test-Path -Path $file -PathType Leaf
              if ($exists) {
                try { $fi = Get-Item -LiteralPath $file; $sizeB = [int]$fi.Length } catch { $sizeB = -1 }
                Write-Output ("[OK] {0}:{1} -> {2} (verified {3} bytes)" -f $ws.Name, $rec.Name, $file, $sizeB)
                $ok++; $okSheet++
              } else {
                Write-Output ("[WARN] Verified missing: {0}:{1} -> {2}" -f $ws.Name, $rec.Name, $file)
                $fail++; $failSheet++
              }
            } else {
              throw "Export failed"
            }
          } catch { $fail++; $failSheet++; Write-Output ("[FAIL] {0}:{1} -> {2}" -f $ws.Name, $rec.Name, $_.Exception.Message) }
        }
        Write-Output ("[SHEET] Summary: {0} ok={1}, fail={2}" -f $ws.Name, $okSheet, $failSheet)
      } catch { $fail++; $failSheet++; Write-Output ("[FAIL_SHEET] {0} -> {1}" -f $ws.Name, $_.Exception.Message) }
    }
    Write-Output ("Summary: exported={0}, failed={1}" -f $ok, $fail)
  } finally {
    $wb.Close($false)
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
  const excelPath = args.excel || 'C\\Users\\jacky\\Downloads\\BRA_006-03_農業経営統計調査（自社内）_テストエビデンス_20251017.xls'
  const sheetName = args.sheet || ''
  const outDir = args.out || args.outDir
  const dryRun = !!args.dryRun

  if (!outDir) {
    console.error('缺少参数: --out <输出目录>')
    process.exit(1)
  }
  if (dryRun) {
    console.log('干跑模式')
    console.log('EXTRACT EXCEL: ' + excelPath)
    console.log('SHEET: ' + sheetName)
    console.log('输出目录: ' + outDir)
    process.exit(0)
  }
  if (!fs.existsSync(excelPath)) {
    console.error('目标EXCEL不存在: ' + excelPath)
    process.exit(1)
  }
  if (!fs.existsSync(outDir)) {
    fs.mkdirSync(outDir, { recursive: true })
  }
  runExcelCOMExtract({ excelPath, sheetName, outDir })
  console.log('完成')
}

main()
