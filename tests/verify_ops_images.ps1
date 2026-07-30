# End-to-end check of "More operations" as a page workshop: real images are added to a REAL
# window through the same code the button calls, one page is rotated, and the assembled set is
# written to a PDF exactly as the window writes it. Then the window is photographed, because a
# grid that shows nothing still passes every assertion about counts.
#
# Settings go to a throw-away folder: the real ones carry window placement and history.
# Usage: powershell -NoProfile -STA -File tests\verify_ops_images.ps1
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

$root = Split-Path -Parent $PSScriptRoot
$exe = Join-Path $root 'dist\iwoHelperDesktop.exe'
if (-not (Test-Path $exe)) { throw "build first: $exe not found" }
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $root 'build\PdfSharp.dll'))
$asm = [System.Reflection.Assembly]::LoadFrom($exe)

$work = Join-Path $env:TEMP ('iwo_verify_ops_' + [Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $work -Force)
$shots = Join-Path $root 'tests\out'
[void](New-Item -ItemType Directory -Path $shots -Force)

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Shot2 {
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint flags);
    [DllImport("kernel32.dll")] public static extern IntPtr GetCurrentProcess();
    [DllImport("kernel32.dll")] public static extern bool TerminateProcess(IntPtr h, uint code);
}
"@

# Leaving this process the normal way crashes AFTER every check has passed: the page grid touches
# Windows.Data.Pdf, and detaching the WinRT runtime at DLL_PROCESS_DETACH dies in a hidden window.
# The app itself ends the same way (see FastExit) - so does the unit runner. Without this the step
# reports a failure that never happened.
function Stop-Now([int]$code) {
    [Console]::Out.Flush()
    [void][Shot2]::TerminateProcess([Shot2]::GetCurrentProcess(), [uint32]$code)
}

function Pump([int]$ms) {
    $end = [DateTime]::UtcNow.AddMilliseconds($ms)
    while ([DateTime]::UtcNow -lt $end) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 25
    }
}

function Field($obj, [string]$name) {
    $flags = [Reflection.BindingFlags]'Instance,NonPublic,Public,FlattenHierarchy'
    $t = $obj.GetType()
    while ($t -ne $null) {
        $f = $t.GetField($name, $flags)
        if ($f -ne $null) { return $f.GetValue($obj) }
        $t = $t.BaseType
    }
    throw "field $name not found"
}

# The parameter is NOT called $args: that name is an automatic variable in PowerShell and the
# call would pass whatever the caller had, not what we hand over.
function Invoke-Hidden($obj, [string]$name, $argv) {
    $flags = [Reflection.BindingFlags]'Instance,NonPublic,Public,FlattenHierarchy'
    $t = $obj.GetType()
    while ($t -ne $null) {
        $m = $t.GetMethod($name, $flags)
        if ($m -ne $null) { return $m.Invoke($obj, $argv) }
        $t = $t.BaseType
    }
    throw "method $name not found"
}

# ---------- sample images ----------
# A wide photo-like JPEG (noise, so a re-encode would show up in the file size) and a tall PNG
# with transparency (it must land on white, not on black).
$jpg = Join-Path $work 'wide.jpg'
$bmp = New-Object System.Drawing.Bitmap 1200, 800
$rnd = New-Object System.Random 7
for ($y = 0; $y -lt 800; $y += 4) {
    for ($x = 0; $x -lt 1200; $x += 4) {
        $bmp.SetPixel($x, $y, [System.Drawing.Color]::FromArgb($rnd.Next(256), $rnd.Next(256), $rnd.Next(256)))
    }
}
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.DrawString('WIDE JPEG', (New-Object System.Drawing.Font 'Arial', 48), [System.Drawing.Brushes]::Black, 40, 40)
$g.Dispose()
$bmp.Save($jpg, [System.Drawing.Imaging.ImageFormat]::Jpeg)
$bmp.Dispose()

$png = Join-Path $work 'tall.png'
$bmp = New-Object System.Drawing.Bitmap 600, 900, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.Clear([System.Drawing.Color]::Transparent)
$g.DrawString('TALL PNG', (New-Object System.Drawing.Font 'Arial', 40), [System.Drawing.Brushes]::Black, 30, 30)
$g.FillRectangle([System.Drawing.Brushes]::Black, 30, 120, 200, 60)
$g.Dispose()
$bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
$bmp.Dispose()

$paths = [Type]::GetType('ExcelMerger.AppPaths, ' + $asm.FullName)
$setRoot = $paths.GetMethod('SetRootForTests', [Reflection.BindingFlags]'Static,NonPublic')
$settingsRoot = [string](Join-Path $work 'settings')
$setRoot.Invoke($null, [object[]]@($settingsRoot))

$loc = [Type]::GetType('ExcelMerger.Loc, ' + $asm.FullName)
$loc.GetMethod('Init', [Reflection.BindingFlags]'Static,Public', $null, @([Type]::GetType('ExcelMerger.Lang, ' + $asm.FullName)), $null).Invoke($null, @(0))

$problems = New-Object System.Collections.ArrayList
$form = New-Object ('ExcelMerger.PdfOpsForm')
try {
    $form.Show()
    Pump 400

    # The very code path of "Add images..." minus the file dialog.
    Invoke-Hidden $form 'AddImages' @(,[string[]]@($jpg, $png))
    $deadline = [DateTime]::UtcNow.AddSeconds(30)
    while ([DateTime]::UtcNow -lt $deadline) {
        Pump 150
        $order = Field $form '_order'
        if ($order.Count -ge 2) { break }
    }
    $order = Field $form '_order'
    if ($order.Count -ne 2) { [void]$problems.Add("two images gave $($order.Count) pages") }

    # Both pages must be A4 sheets, one landscape (wide photo) and one portrait (tall shot).
    $sizes = @()
    foreach ($i in 0..($order.Count - 1)) {
        $ref = $order.Item($i)
        $doc = [PdfSharp.Pdf.IO.PdfReader]::Open($ref.SourcePath, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::InformationOnly)
        $page = $doc.Pages[$ref.PageIndex]
        $sizes += ('{0}x{1}' -f [Math]::Round($page.Width.Point), [Math]::Round($page.Height.Point))
        $doc.Dispose()
    }
    if ($sizes[0] -ne '842x595') { [void]$problems.Add("wide image sheet is $($sizes[0]), expected landscape A4 842x595") }
    if ($sizes[1] -ne '595x842') { [void]$problems.Add("tall image sheet is $($sizes[1]), expected portrait A4 595x842") }

    # Rotate the first page through the grid, as the user does with the tile buttons.
    $grid = Field $form '_grid'
    $grid.SelectIndex(0)
    $grid.RotateSelected(90)
    Pump 200
    if ($order.Item(0).Rotation -ne 90) { [void]$problems.Add('rotation did not reach the page') }

    # Ctrl+Z must take it back: rotation goes into the same history as reordering.
    $undo = Invoke-Hidden $form 'UndoOrder' @()
    Pump 200
    $order = Field $form '_order'
    if ($order.Item(0).Rotation -ne 0) { [void]$problems.Add('Ctrl+Z did not undo the rotation') }
    $grid.SelectIndex(0)
    $grid.RotateSelected(90)
    Pump 200

    # What "Save PDF..." writes, written the same way.
    $merge = [Type]::GetType('ExcelMerger.PdfMergeService, ' + $asm.FullName)
    # Join-Path returns a PSObject wrapper: reflection needs a real string, or the call fails
    # with "cannot convert PSObject to String" - the same trap as with the settings root above.
    $outPdf = [string](Join-Path $work 'assembled.pdf')
    $list = $order.ToList()
    $merge.GetMethod('Merge').Invoke($null, @([object]$list, [object]$outPdf, $null, $null, [object]$false))
    $doc = [PdfSharp.Pdf.IO.PdfReader]::Open($outPdf, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::InformationOnly)
    if ($doc.PageCount -ne 2) { [void]$problems.Add("assembled PDF has $($doc.PageCount) pages") }
    $rotated = $doc.Pages[0].Elements.GetInteger('/Rotate')
    if ($rotated -ne 90) { [void]$problems.Add("rotation lost in the file (/Rotate = $rotated)") }
    $doc.Dispose()

    # A grid that renders nothing still counts pages right, so look at it.
    Pump 900
    $bounds = $form.Bounds
    $img = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $gg = [System.Drawing.Graphics]::FromImage($img)
    $hdc = $gg.GetHdc()
    [void][Shot2]::PrintWindow($form.Handle, $hdc, 0)
    $gg.ReleaseHdc($hdc)
    $gg.Dispose()
    $shot = Join-Path $shots 'ops-images.png'
    $img.Save($shot, [System.Drawing.Imaging.ImageFormat]::Png)
    $img.Dispose()
    Write-Host "screenshot: $shot"
    Write-Host "assembled: $outPdf"
    $form.Close()
} finally {
    $form.Dispose()
    $setRoot.Invoke($null, [object[]]@($null))
}

if ($problems.Count -gt 0) {
    Write-Host ''
    Write-Host 'PROBLEMS:'
    $problems | ForEach-Object { Write-Host " - $_" }
    Stop-Now 1
}
Write-Host 'OK: images become A4 pages, rotation survives undo and reaches the file'
Stop-Now 0
