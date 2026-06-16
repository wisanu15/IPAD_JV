param(
  [string]$ConfigPath = "$PSScriptRoot\rich-menu.json",
  [string]$ImagePath = "$PSScriptRoot\rich-menu.png",
  [switch]$SkipUpload,
  [switch]$NoSetDefault
)

$ErrorActionPreference = "Stop"

function New-RichMenuImage {
  param([string]$Path)

  Add-Type -AssemblyName System.Drawing

  $width = 2500
  $height = 1686
  $bitmap = New-Object System.Drawing.Bitmap $width, $height
  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
  $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

  $bg = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    (New-Object System.Drawing.Rectangle 0,0,$width,$height),
    [System.Drawing.Color]::FromArgb(245,248,255),
    [System.Drawing.Color]::FromArgb(232,240,255),
    45
  )
  $graphics.FillRectangle($bg, 0, 0, $width, $height)

  $titleFont = New-Object System.Drawing.Font "Tahoma", 58, ([System.Drawing.FontStyle]::Bold)
  $subFont = New-Object System.Drawing.Font "Tahoma", 30, ([System.Drawing.FontStyle]::Regular)
  $iconFont = New-Object System.Drawing.Font "Tahoma", 86, ([System.Drawing.FontStyle]::Bold)
  $white = [System.Drawing.Brushes]::White
  $dark = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(28,38,64))
  $muted = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(92,105,130))
  $borderPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(210,220,238)), 4

  $items = @(
    @{ X=0;    Y=0;   W=833; H=843; Icon="01"; Title="STATUS"; Sub="Check iPad info" ; Color=[System.Drawing.Color]::FromArgb(79,126,247) },
    @{ X=833;  Y=0;   W=833; H=843; Icon="02"; Title="REPORT"; Sub="iPad Pen Case Cable" ; Color=[System.Drawing.Color]::FromArgb(245,158,11) },
    @{ X=1666; Y=0;   W=834; H=843; Icon="03"; Title="BORROW"; Sub="Borrow and return" ; Color=[System.Drawing.Color]::FromArgb(16,185,129) },
    @{ X=0;    Y=843; W=833; H=843; Icon="04"; Title="APP REQUEST"; Sub="Request learning app" ; Color=[System.Drawing.Color]::FromArgb(139,92,246) },
    @{ X=833;  Y=843; W=833; H=843; Icon="05"; Title="GUIDE"; Sub="How to use" ; Color=[System.Drawing.Color]::FromArgb(6,182,212) },
    @{ X=1666; Y=843; W=834; H=843; Icon="06"; Title="OPEN APP"; Sub="iPad JV home" ; Color=[System.Drawing.Color]::FromArgb(236,72,153) }
  )

  foreach ($item in $items) {
    $rect = New-Object System.Drawing.Rectangle $item.X, $item.Y, $item.W, $item.H
    $graphics.FillRectangle($white, $rect)
    $graphics.DrawRectangle($borderPen, $rect)

    $accent = New-Object System.Drawing.SolidBrush $item.Color
    $graphics.FillRectangle($accent, $item.X, $item.Y, $item.W, 18)

    $center = New-Object System.Drawing.StringFormat
    $center.Alignment = [System.Drawing.StringAlignment]::Center
    $center.LineAlignment = [System.Drawing.StringAlignment]::Center

    $graphics.DrawString($item.Icon, $iconFont, $dark, (New-Object System.Drawing.RectangleF $item.X, ($item.Y + 155), $item.W, 130), $center)
    $graphics.DrawString($item.Title, $titleFont, $dark, (New-Object System.Drawing.RectangleF $item.X, ($item.Y + 355), $item.W, 90), $center)
    $graphics.DrawString($item.Sub, $subFont, $muted, (New-Object System.Drawing.RectangleF $item.X, ($item.Y + 465), $item.W, 70), $center)
  }

  $outDir = Split-Path -Parent $Path
  if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
  $bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)

  $graphics.Dispose()
  $bitmap.Dispose()
}

New-RichMenuImage -Path $ImagePath
Write-Host "Created image: $ImagePath"

if ($SkipUpload) {
  Write-Host "Skipped upload."
  exit 0
}

$token = [Environment]::GetEnvironmentVariable("LINE_CHANNEL_ACCESS_TOKEN", "User")
if (!$token) { $token = [Environment]::GetEnvironmentVariable("LINE_CHANNEL_ACCESS_TOKEN", "Process") }
if (!$token) { throw "Set LINE_CHANNEL_ACCESS_TOKEN first, or run with -SkipUpload to only create the image." }

$headers = @{ Authorization = "Bearer $token" }
$json = Get-Content -Path $ConfigPath -Raw -Encoding UTF8
$created = Invoke-RestMethod -Method Post -Uri "https://api.line.me/v2/bot/richmenu" -Headers $headers -ContentType "application/json" -Body $json
$richMenuId = $created.richMenuId
Write-Host "Created rich menu: $richMenuId"

Invoke-RestMethod -Method Post -Uri "https://api-data.line.me/v2/bot/richmenu/$richMenuId/content" -Headers $headers -ContentType "image/png" -InFile $ImagePath | Out-Null
Write-Host "Uploaded image."

if (!$NoSetDefault) {
  Invoke-RestMethod -Method Post -Uri "https://api.line.me/v2/bot/user/all/richmenu/$richMenuId" -Headers $headers | Out-Null
  Write-Host "Set as default rich menu."
}

Write-Host "Done."
