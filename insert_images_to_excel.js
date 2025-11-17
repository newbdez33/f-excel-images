const fs = require('fs')
const path = require('path')
const os = require('os')
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

function getImages(dir) {
  const exts = new Set(['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff'])
  const out = []
  function walk(d) {
    const entries = fs.readdirSync(d, { withFileTypes: true })
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }))
    for (const e of entries) {
      const fp = path.join(d, e.name)
      if (e.isDirectory()) {
        walk(fp)
      } else if (e.isFile()) {
        const ext = path.extname(e.name).toLowerCase()
        if (exts.has(ext)) out.push(fp)
      }
    }
  }
  function numKey(p) {
    const bn = path.basename(p, path.extname(p))
    if (/^\d+$/.test(bn)) return parseInt(bn, 10)
    const m = bn.match(/\d+/)
    return m ? parseInt(m[0], 10) : Number.POSITIVE_INFINITY
  }
  walk(dir)
  return out.sort((a, b) => {
    const ta = (path.relative(dir, path.dirname(a)).split(/[\\/]/)[0] || '').toLowerCase()
    const tb = (path.relative(dir, path.dirname(b)).split(/[\\/]/)[0] || '').toLowerCase()
    if (ta !== tb) return ta.localeCompare(tb, undefined, { numeric: true, sensitivity: 'base' })
    const na = numKey(a), nb = numKey(b)
    if (na !== nb) return na - nb
    const ba = path.basename(a), bb = path.basename(b)
    return ba.localeCompare(bb, undefined, { numeric: true, sensitivity: 'base' })
  })
}

function runExcelCOM({ excelPath, sheetName, imageDir, templateRow, imageCol, recordCol, files }) {
  const list = files && files.length ? files : getImages(imageDir)
  const tmpList = path.join(os.tmpdir(), `excel_fuck_image_list_${Date.now()}.txt`)
  fs.writeFileSync(tmpList, list.join('\n'), 'utf8')
  const ps = `
  $ErrorActionPreference = 'Stop'
  $excelPath = '${excelPath.replace(/'/g, "''")}'
  $sheetName = '${sheetName.replace(/'/g, "''")}'
  $imageDir = '${imageDir.replace(/'/g, "''")}'
  $templateRow = ${templateRow}
  $imageCol = ${imageCol}
  $recordCol = ${recordCol ?? 0}
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($excelPath)
  try {
    $ws = $null
    foreach ($w in $wb.Worksheets) { if ($w.Name -eq $sheetName) { $ws = $w; break } }
    if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
    $usedLast = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    if ($usedLast -ge 4) { $ws.Rows.Item("4:$usedLast").Delete(); Write-Output ("Cleared rows 4:{0}" -f $usedLast) }
    $lastRow = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
    $firstDestRow = $lastRow + 1
    $deletedShapes = 0
    foreach ($s in $ws.Shapes) { try { if ($s.TopLeftCell.Row -ge 4) { $s.Delete(); $deletedShapes++ } } catch {} }
    Write-Output ("Cleared shapes rows>=4: {0}" -f $deletedShapes)
    $listPath = '${tmpList.replace(/'/g, "''")}'
    $files = Get-Content -LiteralPath $listPath -Encoding UTF8
    Add-Type -AssemblyName System.Drawing
    $countTotal = 0; $countOk = 0; $countFail = 0
    foreach ($f in $files) {
      $destRow = $lastRow + 1
      try {
        $ws.Rows.Item($templateRow).Copy($ws.Rows.Item($destRow))
        $cell = $ws.Cells.Item($destRow, $imageCol)
        if ($cell.MergeCells) { try { $cell.MergeArea.UnMerge() } catch {} ; $cell = $ws.Cells.Item($destRow, $imageCol) }
        $cellLeft = [double]$cell.Left; $cellTop = [double]$cell.Top; $cellW = [double]$cell.Width; $cellH = [double]$cell.Height
        if ($cellW -le 0 -or $cellH -le 0) { throw "Target cell has zero size" }
        $noCol = [int]($imageCol - 1)
        if ($noCol -ge 1) {
          $noCell = $ws.Cells.Item($destRow, $noCol)
          if ($destRow -ne $firstDestRow) {
            $prevVal = $null
            try { $prevVal = $ws.Cells.Item($destRow - 1, $noCol).Value2 } catch {}
            if ($prevVal -ne $null -and $prevVal -ne '') {
              try { $noCell.Value2 = [double]$prevVal + 1 } catch { $noCell.Value2 = $prevVal }
            } else {
              $tmplVal = $null
              try { $tmplVal = $ws.Cells.Item($templateRow, $noCol).Value2 } catch {}
              if ($tmplVal -ne $null -and $tmplVal -ne '') {
                try { $noCell.Value2 = [double]$tmplVal + 1 } catch { $noCell.Value2 = $tmplVal }
              }
            }
          }
        }
        $img = [System.Drawing.Image]::FromFile($f)
        $imgW = [double]$img.Width; $imgH = [double]$img.Height
        $img.Dispose()
        if ($imgW -le 0 -or $imgH -le 0) { throw "Image has zero dimension" }
        if (($imgW -lt $cellW) -and ($imgH -lt $cellH)) {
          $newW = $imgW; $newH = $imgH
        } else {
          $scaleW = $cellW / $imgW; $scaleH = $cellH / $imgH
          if ($scaleW -lt $scaleH) {
            $newW = $cellW; $newH = $imgH * $scaleW
          } else {
            $newH = $cellH; $newW = $imgW * $scaleH
          }
        }
        $left = $cellLeft
        $top = $cellTop
        $shapesBefore = $ws.Shapes.Count
        $shape = $ws.Shapes.AddPicture($f, $false, $true, [double]$left, [double]$top, [double]$newW, [double]$newH)
        try { $shape.LockAspectRatio = $true } catch {}
        try { $shape.Placement = 2 } catch {}
        if ($recordCol -gt 0) { $ws.Cells.Item($destRow, $recordCol).Value2 = [System.IO.Path]::GetFileName($f) }
        $shapesAfter = $ws.Shapes.Count
        if ($shape -ne $null -and $shapesAfter -gt $shapesBefore) {
          Write-Output "[OK] $($f) -> Row $destRow, Col $imageCol"
          $countOk++
        } else {
          Write-Output "[FAIL] $($f) -> Shape not added"
          $countFail++
        }
        $lastRow = $destRow
      } catch {
        Write-Output "[FAIL] $($f) -> $($_.Exception.Message)"
        $countFail++
      }
      $countTotal++
    }
    Write-Output ("Summary: total={0}, ok={1}, fail={2}" -f $countTotal, $countOk, $countFail)
    $wb.Save()
  } finally {
    $wb.Close($true)
    $excel.Quit()
  }
  `
  const encoded = Buffer.from(ps, 'utf16le').toString('base64')
  const r = spawnSync('powershell.exe', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded], { stdio: 'inherit' })
  try { fs.unlinkSync(tmpList) } catch {}
  if (r.error) throw r.error
  if (r.status !== 0) throw new Error('PowerShell exited with code ' + r.status)
}


function main() {
  const args = parseArgs(process.argv)
  const excelPath = args.excel || 'C:\\Users\\jacky\\Downloads\\BRA_006-03_農業経営統計調査（自社内）_テストエビデンス_20251017.xls'
  const sheetName = args.sheet || '農政局'
  const imageDir = args.dir
  const templateRow = parseInt(args.templateRow || args.copyRow || '1', 10)
  const imageCol = parseInt(args.imageCol || '1', 10)
  const recordCol = args.recordCol ? parseInt(args.recordCol, 10) : undefined
  const dryRun = !!args.dryRun
  

  if (!imageDir) {
    console.error('缺少参数: --dir <图片目录>')
    process.exit(1)
  }
  if (!fs.existsSync(imageDir)) {
    console.error('图片目录不存在: ' + imageDir)
    process.exit(1)
  }
  const imgs = getImages(imageDir)
  if (dryRun) {
    console.log('干跑模式')
    console.log('EXCEL: ' + excelPath)
    console.log('SHEET: ' + sheetName)
    console.log('模板行: ' + templateRow)
    console.log('图片列: ' + imageCol)
    console.log('记录文件名列: ' + (recordCol || '无'))
    console.log('图片数量: ' + imgs.length)
    imgs.forEach(f => console.log(f))
    process.exit(0)
  }
  if (!fs.existsSync(excelPath)) {
    console.error('目标EXCEL不存在: ' + excelPath)
    process.exit(1)
  }
  if (imgs.length === 0) {
    console.error('目录中未找到图片: ' + imageDir)
    process.exit(1)
  }
  const backupPath = createBackup(excelPath)
  console.log('已备份: ' + backupPath)
  runExcelCOM({ excelPath, sheetName, imageDir, templateRow, imageCol, recordCol, files: imgs })
  console.log('完成')
}

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

main()
