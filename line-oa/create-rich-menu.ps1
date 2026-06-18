param(
  [string]$ConfigPath = "$PSScriptRoot\rich-menu.json",
  [string]$ImagePath  = "$PSScriptRoot\rich-menu.png",
  [switch]$SkipUpload,
  [switch]$NoSetDefault
)
$ErrorActionPreference = "Stop"
function U { param([string]$s); [regex]::Unescape($s) }
function E { param([int]$cp); [char]::ConvertFromUtf32($cp) }

function New-RichMenuImage { param([string]$Path)
  Add-Type -AssemblyName System.Drawing
  $W=2500; $H=1686
  $bmp=New-Object System.Drawing.Bitmap $W,$H
  $g=[System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode      =[System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.TextRenderingHint  =[System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
  $g.CompositingQuality =[System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $g.InterpolationMode  =[System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

  # Background: #0F172A (sidebar color)
  $g.Clear([System.Drawing.Color]::FromArgb(15,23,42))

  # Blue gradient overlay (matches sidebar::before rgba(79,126,247,0.14))
  $rect0=New-Object System.Drawing.Rectangle 0,0,$W,$H
  $bgBrush=New-Object System.Drawing.Drawing2D.LinearGradientBrush(
    $rect0,
    [System.Drawing.Color]::FromArgb(36,79,126,247),
    [System.Drawing.Color]::FromArgb(0,79,126,247),
    [System.Drawing.Drawing2D.LinearGradientMode]::ForwardDiagonal)
  $g.FillRectangle($bgBrush,0,0,$W,$H)

  # Ambient orbs
  function Draw-Orb { param($gfx,[int]$x,[int]$y,[int]$sz,[int]$r,[int]$gr,[int]$gb,[int]$a)
    $op=New-Object System.Drawing.Drawing2D.GraphicsPath
    $op.AddEllipse($x,$y,$sz,$sz)
    $ob=New-Object System.Drawing.Drawing2D.PathGradientBrush($op)
    $ob.CenterColor=[System.Drawing.Color]::FromArgb($a,$r,$gr,$gb)
    $ob.SurroundColors=@([System.Drawing.Color]::FromArgb(0,$r,$gr,$gb))
    $gfx.FillPath($ob,$op)
    $ob.Dispose(); $op.Dispose()
  }
  Draw-Orb $g  1800 -300 1400  79 126 247 45
  Draw-Orb $g  -300 1100 1200  16 185 129 35
  Draw-Orb $g   600  300  800 139  92 246 25

  # Cards - app palette colors
  # primary:#4F7EF7  warning:#F59E0B  purple:#8B5CF6  cyan:#06B6D4  success:#10B981  danger:#EF4444
  $cards=@(
    @{ X=0;    Y=0;   CW=833; CH=843; Emoji=(E 0x1F50D); T=(U 'ตรวจสถานะ');  S=(U 'ดูข้อมูลไอแพด'); AR=79;  AG=126; AB=247 },
    @{ X=833;  Y=0;   CW=833; CH=843; Emoji=(U '⚠');     T=(U 'แจ้งปัญหา'); S=(U 'ไอแพดหาย/เสีย'); AR=245; AG=158; AB=11  },
    @{ X=1666; Y=0;   CW=834; CH=843; Emoji=(E 0x1F4F1); T=(U 'ยืม / คืน');  S=(U 'ยืม-คืนไอแพด'); AR=139; AG=92;  AB=246 },
    @{ X=0;    Y=843; CW=833; CH=843; Emoji=(E 0x1F4F2); T=(U 'ขอแอพ');      S=(U 'ขอติดตั้งแอพ'); AR=6;   AG=182; AB=212 },
    @{ X=833;  Y=843; CW=833; CH=843; Emoji=(E 0x1F4D6); T=(U 'คู่มือ');      S=(U 'วิธีใช้งาน');   AR=16;  AG=185; AB=129 },
    @{ X=1666; Y=843; CW=834; CH=843; Emoji=(E 0x1F3E0); T=(U 'เปิดแอป');    S="iPad JV";            AR=239; AG=68;  AB=68  }
  )

  # Fonts - larger than before
  $fT =New-Object System.Drawing.Font "Leelawadee UI",80,([System.Drawing.FontStyle]::Bold)
  $fS =New-Object System.Drawing.Font "Leelawadee UI",36,([System.Drawing.FontStyle]::Regular)
  $fE =New-Object System.Drawing.Font "Segoe UI Emoji",110,([System.Drawing.FontStyle]::Regular)
  $fSy=New-Object System.Drawing.Font "Segoe UI Symbol",110,([System.Drawing.FontStyle]::Regular)
  $sfc=New-Object System.Drawing.StringFormat
  $sfc.Alignment    =[System.Drawing.StringAlignment]::Center
  $sfc.LineAlignment=[System.Drawing.StringAlignment]::Center
  $brW=[System.Drawing.Brushes]::White
  $brM=New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(200,255,255,255))

  $m=10

  foreach($c in $cards){
    [int]$cx=$c.X; [int]$cy=$c.Y; [int]$cw=$c.CW; [int]$ch=$c.CH
    [int]$ar=$c.AR; [int]$ag=$c.AG; [int]$ab=$c.AB
    [int]$rx=$cx+$m; [int]$ry=$cy+$m; [int]$rw=$cw-$m*2; [int]$rh=$ch-$m*2

    # Card fill - slightly lighter than background
    $brCard=New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30,255,255,255))
    $g.FillRectangle($brCard,$rx,$ry,$rw,$rh)

    # Card border
    $penB=New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(40,255,255,255)),1.5
    $g.DrawRectangle($penB,$rx,$ry,$rw,$rh)

    # Top accent strip (thicker: 24px)
    $brAcc=New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($ar,$ag,$ab))
    $g.FillRectangle($brAcc,$rx,$ry,$rw,24)

    # Gradient fade below strip
    [int]$stripY=$ry+24; [int]$stripH=120
    $rectStrip=New-Object System.Drawing.Rectangle $rx,$stripY,$rw,$stripH
    $gradS=New-Object System.Drawing.Drawing2D.LinearGradientBrush(
      $rectStrip,
      [System.Drawing.Color]::FromArgb(50,$ar,$ag,$ab),
      [System.Drawing.Color]::FromArgb(0,$ar,$ag,$ab),
      [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
    $g.FillRectangle($gradS,$rx,$stripY,$rw,$stripH)

    # Icon glow halo
    [int]$hSz=300; [int]$hX=$cx+($cw-$hSz)/2; [int]$hY=$cy+75
    $hp=New-Object System.Drawing.Drawing2D.GraphicsPath
    $hp.AddEllipse($hX,$hY,$hSz,$hSz)
    $hb=New-Object System.Drawing.Drawing2D.PathGradientBrush($hp)
    $hb.CenterColor=[System.Drawing.Color]::FromArgb(65,$ar,$ag,$ab)
    $hb.SurroundColors=@([System.Drawing.Color]::FromArgb(0,$ar,$ag,$ab))
    $g.FillPath($hb,$hp)
    $hb.Dispose(); $hp.Dispose()

    # Icon circle
    [int]$iSz=210; [int]$iX=$cx+($cw-$iSz)/2; [int]$iY=$cy+115
    $brIC=New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(60,$ar,$ag,$ab))
    $g.FillEllipse($brIC,$iX,$iY,$iSz,$iSz)
    $penIC=New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(100,$ar,$ag,$ab)),2
    $g.DrawEllipse($penIC,$iX,$iY,$iSz,$iSz)

    # Emoji icon
    $ir=New-Object System.Drawing.RectangleF([float]$iX,[float]$iY,[float]$iSz,[float]$iSz)
    try   { $g.DrawString($c.Emoji,$fE,$brW,$ir,$sfc) }
    catch { $g.DrawString($c.Emoji,$fSy,$brW,$ir,$sfc) }

    # Title - 80pt bold
    [int]$tY=$cy+370
    $tr=New-Object System.Drawing.RectangleF([float]$cx,[float]$tY,[float]$cw,140.0)
    $g.DrawString($c.T,$fT,$brW,$tr,$sfc)

    # Subtitle (36pt)
    [int]$sY=$cy+530
    $sr=New-Object System.Drawing.RectangleF([float]$cx,[float]$sY,[float]$cw,80.0)
    $g.DrawString($c.S,$fS,$brM,$sr,$sfc)
  }

  # Subtle dot grid
  $dotBr=New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(6,255,255,255))
  for($dy=30; $dy -lt $H; $dy+=70){
    for($dx=30; $dx -lt $W; $dx+=70){
      $g.FillEllipse($dotBr,$dx,$dy,4,4)
    }
  }

  $outDir=Split-Path -Parent $Path
  if($outDir -and !(Test-Path $outDir)){ New-Item -ItemType Directory -Path $outDir | Out-Null }
  $bmp.Save($Path,[System.Drawing.Imaging.ImageFormat]::Png)
  foreach($o in @($g,$bmp,$bgBrush,$fT,$fS,$fE,$fSy,$sfc,$brM)){try{$o.Dispose()}catch{}}
}

New-RichMenuImage -Path $ImagePath
Write-Host "Created image: $ImagePath"
if($SkipUpload){ Write-Host "Skipped upload."; exit 0 }
$token=[Environment]::GetEnvironmentVariable("LINE_CHANNEL_ACCESS_TOKEN","User")
if(!$token){$token=[Environment]::GetEnvironmentVariable("LINE_CHANNEL_ACCESS_TOKEN","Process")}
if(!$token){ throw "Set LINE_CHANNEL_ACCESS_TOKEN first." }
$headers=@{Authorization="Bearer $token"}
$json=Get-Content -Path $ConfigPath -Raw -Encoding UTF8
$created=Invoke-RestMethod -Method Post -Uri "https://api.line.me/v2/bot/richmenu" -Headers $headers -ContentType "application/json" -Body $json
$richMenuId=$created.richMenuId
Write-Host "Created rich menu: $richMenuId"
Invoke-RestMethod -Method Post -Uri "https://api-data.line.me/v2/bot/richmenu/$richMenuId/content" -Headers $headers -ContentType "image/png" -InFile $ImagePath | Out-Null
Write-Host "Uploaded image."
if(!$NoSetDefault){
  Invoke-RestMethod -Method Post -Uri "https://api.line.me/v2/bot/user/all/richmenu/$richMenuId" -Headers $headers | Out-Null
  Write-Host "Set as default rich menu."
}
Write-Host "Done."

