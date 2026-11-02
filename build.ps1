# Compila o Conversor PDF com o compilador C# que já vem no Windows (.NET Framework 4.x)
param([string]$Saida = "$PSScriptRoot\bin\Conversor PDF.exe")

$ErrorActionPreference = 'Stop'
$fw = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319"
$csc = "$fw\csc.exe"
New-Item -ItemType Directory -Force (Split-Path $Saida) | Out-Null

# Ícone (gerado na hora: folha branca com faixa vermelha "PDF")
Add-Type -AssemblyName System.Drawing
$ico = Join-Path (Split-Path $Saida) "icone.ico"
$tamanhos = 256, 48, 32, 16
$pngs = foreach ($t in $tamanhos) {
    $b = New-Object System.Drawing.Bitmap $t, $t
    $g = [System.Drawing.Graphics]::FromImage($b)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'AntiAliasGridFit'
    $s = $t / 256.0
    $folha = New-Object System.Drawing.Drawing2D.GraphicsPath
    $pts = @(
        (New-Object System.Drawing.PointF (40 * $s), (12 * $s)),
        (New-Object System.Drawing.PointF (170 * $s), (12 * $s)),
        (New-Object System.Drawing.PointF (222 * $s), (64 * $s)),
        (New-Object System.Drawing.PointF (222 * $s), (244 * $s)),
        (New-Object System.Drawing.PointF (40 * $s), (244 * $s)))
    $folha.AddPolygon($pts)
    $g.FillPath([System.Drawing.Brushes]::White, $folha)
    $g.DrawPath((New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(150, 150, 160)), ([Math]::Max(1, 6 * $s))), $folha)
    $dobra = @(
        (New-Object System.Drawing.PointF (170 * $s), (12 * $s)),
        (New-Object System.Drawing.PointF (170 * $s), (64 * $s)),
        (New-Object System.Drawing.PointF (222 * $s), (64 * $s)))
    $g.FillPolygon((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(215, 218, 225))), $dobra)
    $faixa = New-Object System.Drawing.RectangleF (8 * $s), (128 * $s), (200 * $s), (84 * $s)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(217, 48, 37))), $faixa)
    if ($t -ge 32) {
        $f = New-Object System.Drawing.Font 'Segoe UI', ([float](54 * $s)), ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
        $g.DrawString('PDF', $f, [System.Drawing.Brushes]::White, $faixa, $sf)
    }
    $g.Dispose()
    $ms = New-Object System.IO.MemoryStream
    $b.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $b.Dispose()
    , $ms.ToArray()
}
$fs = [System.IO.File]::Create($ico)
$bw = New-Object System.IO.BinaryWriter $fs
$bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$tamanhos.Count)
$offset = 6 + 16 * $tamanhos.Count
for ($i = 0; $i -lt $tamanhos.Count; $i++) {
    $t = $tamanhos[$i]
    $bw.Write([byte]($t % 256)); $bw.Write([byte]($t % 256)); $bw.Write([byte]0); $bw.Write([byte]0)
    $bw.Write([uint16]1); $bw.Write([uint16]32)
    $bw.Write([uint32]$pngs[$i].Length); $bw.Write([uint32]$offset)
    $offset += $pngs[$i].Length
}
foreach ($p in $pngs) { $bw.Write($p) }
$bw.Close()

$refs = @('System.Windows.Forms.dll', 'System.Drawing.dll', 'System.IO.Compression.dll', 'System.IO.Compression.FileSystem.dll',
          'System.Xml.dll', 'System.Xml.Linq.dll', 'System.Core.dll', 'Microsoft.CSharp.dll', 'System.Xaml.dll',
          "$fw\WPF\PresentationCore.dll", "$fw\WPF\WindowsBase.dll") | ForEach-Object { "/r:$_" }
$fontes = Get-ChildItem "$PSScriptRoot\src\*.cs" | ForEach-Object { $_.FullName }
& $csc /nologo /codepage:65001 /optimize+ /target:winexe "/win32icon:$ico" "/out:$Saida" $refs $fontes
if ($LASTEXITCODE -ne 0) { throw "Falha na compilação" }
"OK: $Saida"
