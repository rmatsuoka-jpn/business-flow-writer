#Requires -Version 5.1
<#
.SYNOPSIS
    フロー定義JSON から BPMN準拠の業務フロー図を Excel (.xlsx) に描画する

.DESCRIPTION
    business-flow-writer スキルの描画エンジン。
    Excel COM でオートシェイプを配置するため、生成後の図は手作業で
    自由にブラッシュアップできる（図形のドラッグ・テキスト編集・線の付け替え）。

    定義JSONのスキーマは references/notation.md を参照。

.PARAMETER Definition
    フロー定義JSONのパス

.PARAMETER Output
    出力する .xlsx のパス

.PARAMETER ExportPdf
    検証用に同名の PDF も出力する

.PARAMETER Visible
    Excel を画面に表示しながら実行する（デバッグ用）

.EXAMPLE
    .\generate-flow.ps1 -Definition flow.json -Output flow.xlsx -ExportPdf
#>
param(
    [Parameter(Mandatory)][string]$Definition,
    [Parameter(Mandatory)][string]$Output,
    [switch]$ExportPdf,
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

# ---- レイアウト定数（pt） ----
$DiagLeft     = 20      # 図全体の左端
$TopY         = 45      # プール開始Y（タイトルの下）
$PoolStripW   = 22      # プール名の縦書き帯
$LaneStripW   = 22      # レーン名の縦書き帯
$ColWidth     = 105     # 1列（時系列1ステップ）の幅
$RowBaseH     = 70      # レーンの基本高さ（row 0 のみの場合）
$RowStepH     = 65      # row が1つ増えるごとの追加高さ
$PoolGap      = 30      # プール間の隙間（既定。定義JSONの poolGap で上書き可）
$FontName     = "Meiryo UI"

# ---- MSO定数 ----
$msoShapeRectangle        = 1
$msoShapeDiamond          = 4
$msoShapeRoundedRectangle = 5
$msoShapeOval             = 9
$msoShapeCan              = 13
$msoShapeLeftBracket      = 29
$msoShapeFlowchartDocument = 67
$msoShapeFoldedCorner     = 16
$msoConnectorStraight     = 1
$msoConnectorElbow        = 2
$msoLineDash              = 4
$msoArrowheadTriangle     = 2
$msoTextOrientationVerticalFarEast = 4

$SiteMap = @{ top = 1; left = 2; bottom = 3; right = 4 }

# ---- 定義読み込み ----
# 相対パスは PowerShell のカレント($PWD)基準で解決する。
# [System.IO.Path]::GetFullPath は .NET の作業ディレクトリ基準になり、
# PowerShell の cd と同期しないため（対話セッションで直接実行すると誤解決する）使わない。
$defPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Definition)
$outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Output)
$def = Get-Content -Raw -Encoding UTF8 $defPath | ConvertFrom-Json

# プール間隔の全体上書き（pt）。開始プール（住民等）を離して見せたい時などに使う
if ($null -ne $def.poolGap) { $PoolGap = [double]$def.poolGap }

$outDir = Split-Path $outPath -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

# ---- レーン形状の計算 ----
# レーンごとの最大 row から高さを決める
$laneRows = @{}
foreach ($ln in $def.lanes) { $laneRows[$ln.id] = 0 }
foreach ($n in $def.nodes) {
    $r = if ($n.row) { [int]$n.row } else { 0 }
    if (-not $laneRows.ContainsKey($n.lane)) { throw "ノード '$($n.id)' のレーン '$($n.lane)' が lanes に存在しません" }
    if ($r -gt $laneRows[$n.lane]) { $laneRows[$n.lane] = $r }
}

$laneTop = @{}; $laneH = @{}
$poolBounds = [ordered]@{}   # pool名 → @{Top;Bottom;HasLaneNames}
$y = $TopY
$prevPool = $null
foreach ($ln in $def.lanes) {
    if ($null -ne $prevPool -and $ln.pool -ne $prevPool) {
        # プール境界の隙間。プール先頭レーンの gapBefore で個別指定できる
        $gap = if ($null -ne $ln.gapBefore) { [double]$ln.gapBefore } else { $PoolGap }
        $y += $gap
    }
    if (-not $poolBounds.Contains($ln.pool)) {
        $poolBounds[$ln.pool] = @{ Top = $y; Bottom = $y; HasLaneNames = $false }
    }
    $h = $RowBaseH + $laneRows[$ln.id] * $RowStepH
    if ($ln.height) { $h = [double]$ln.height }
    $laneTop[$ln.id] = $y
    $laneH[$ln.id]   = $h
    if ("$($ln.name)".Trim() -ne "") { $poolBounds[$ln.pool].HasLaneNames = $true }
    $y += $h
    $poolBounds[$ln.pool].Bottom = $y
    $prevPool = $ln.pool
}
$diagBottom = $y

$ncols = 1
foreach ($n in $def.nodes) { if ([int]$n.col -gt $ncols) { $ncols = [int]$n.col } }
$NodeAreaLeft = $DiagLeft + $PoolStripW + $LaneStripW
$DiagRight    = $NodeAreaLeft + $ncols * $ColWidth + 15

function Get-NodeX([object]$n) { return $NodeAreaLeft + ([int]$n.col - 0.5) * $ColWidth }
function Get-NodeY([object]$n) {
    $r = if ($n.row) { [int]$n.row } else { 0 }
    return $laneTop[$n.lane] + 35 + $r * $RowStepH
}

# ---- Excel起動 ----
$xl = New-Object -ComObject Excel.Application
$xl.Visible = [bool]$Visible
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "業務フロー"
    $xl.ActiveWindow.DisplayGridlines = $false

    # ---- 共通ヘルパー ----
    function Set-ShapeText($shp, [string]$text, [double]$size = 10, [int]$align = 2, [bool]$bold = $false) {
        $tf = $shp.TextFrame2
        $tr = $tf.TextRange
        $tr.Text = $text.Replace("`r`n", "`r").Replace("`n", "`r")
        $f = $tr.Font
        $f.Name = $FontName; $f.NameFarEast = $FontName
        $f.Size = $size
        $f.Fill.ForeColor.RGB = 0
        $f.Bold = if ($bold) { -1 } else { 0 }
        $tr.ParagraphFormat.Alignment = $align
        $tf.VerticalAnchor = 3
        $tf.MarginLeft = 2; $tf.MarginRight = 2; $tf.MarginTop = 1; $tf.MarginBottom = 1
        $tf.WordWrap = -1
    }
    function Set-BlackLine($shp, [double]$weight = 1.0) {
        $shp.Line.Visible = -1
        $shp.Line.ForeColor.RGB = 0
        $shp.Line.Weight = $weight
    }
    function New-Box([int]$type, [double]$l, [double]$t, [double]$w, [double]$h, [bool]$whiteFill = $true) {
        $shp = $ws.Shapes.AddShape($type, $l, $t, $w, $h)
        if ($whiteFill) { $shp.Fill.ForeColor.RGB = 16777215; $shp.Fill.Visible = -1 }
        else            { $shp.Fill.Visible = 0 }
        Set-BlackLine $shp 1.0
        $shp.Shadow.Visible = 0
        return $shp
    }
    function New-Label([double]$l, [double]$t, [double]$w, [double]$h, [string]$text, [double]$size = 8.5, [int]$align = 2) {
        $shp = $ws.Shapes.AddTextbox(1, $l, $t, $w, $h)
        $shp.Fill.Visible = 0
        $shp.Line.Visible = 0
        Set-ShapeText $shp $text $size $align
        return $shp
    }

    # ---- タイトル ----
    $titleShp = New-Label -l $DiagLeft -t 12 -w 500 -h 24 -text "【業務フロー】$($def.title)" -size 12 -align 1
    $titleShp.TextFrame2.TextRange.Font.Bold = -1

    # ---- プール・レーンの枠 ----
    foreach ($poolName in $poolBounds.Keys) {
        $pb = $poolBounds[$poolName]
        $pTop = $pb.Top; $pH = $pb.Bottom - $pb.Top
        # 外枠
        New-Box $msoShapeRectangle $DiagLeft $pTop ($DiagRight - $DiagLeft) $pH $false | Out-Null
        # プール名帯（縦書き）
        $strip = New-Box $msoShapeRectangle $DiagLeft $pTop $PoolStripW $pH $false
        Set-ShapeText $strip $poolName 10 2
        $strip.TextFrame2.Orientation = $msoTextOrientationVerticalFarEast
        # レーン帯（レーン名があるプールのみ）
        if ($pb.HasLaneNames) {
            foreach ($ln in $def.lanes) {
                if ($ln.pool -ne $poolName) { continue }
                $lstrip = New-Box $msoShapeRectangle ($DiagLeft + $PoolStripW) $laneTop[$ln.id] $LaneStripW $laneH[$ln.id] $false
                Set-ShapeText $lstrip "$($ln.name)" 9 2
                $lstrip.TextFrame2.Orientation = $msoTextOrientationVerticalFarEast
                # レーン境界の横線（外枠と重なる分は気にしない）
                New-Box $msoShapeRectangle ($DiagLeft + $PoolStripW) $laneTop[$ln.id] ($DiagRight - $DiagLeft - $PoolStripW) $laneH[$ln.id] $false | Out-Null
            }
        }
    }

    # ---- ノード ----
    $nodeShapes = @{}
    foreach ($n in $def.nodes) {
        $x = Get-NodeX $n
        $cy = Get-NodeY $n
        # 位置の微調整（pt単位、省略可）
        if ($null -ne $n.dx) { $x  += [double]$n.dx }
        if ($null -ne $n.dy) { $cy += [double]$n.dy }
        $shp = $null
        switch ($n.type) {
            "task" {
                $w = if ($n.width) { [double]$n.width } else { 88 }
                $labelLines = ("$($n.label)" -split "\r?\n").Count
                $h = if ($n.height) { [double]$n.height } else { 40 + ($labelLines - 1) * 13 }
                $shp = New-Box $msoShapeRoundedRectangle ($x - $w/2) ($cy - $h/2) $w $h
                $shp.Adjustments.Item(1) = 0.12
                if ($n.no) {
                    Set-ShapeText $shp "$($n.no)`r$($n.label)" 10 2
                    $tr = $shp.TextFrame2.TextRange
                    $p1 = $tr.Paragraphs(1, 1)
                    $p1.Font.Size = 7.5
                    $p1.ParagraphFormat.Alignment = 1
                } else {
                    Set-ShapeText $shp "$($n.label)" 10 2
                }
            }
            "gateway" {
                $w = 40; $h = 40
                $shp = New-Box $msoShapeDiamond ($x - $w/2) ($cy - $h/2) $w $h
                Set-ShapeText $shp "×" 13 2 $true
                if ($n.label) {
                    # 既定はひし形の上・中央。レーン上端からはみ出す場合は左下に置く
                    # （下の分岐線・右の分岐ラベルと重ならない位置）
                    if (($cy - $h/2 - 20) -ge ($laneTop[$n.lane] + 2)) {
                        New-Label -l ($x - 60) -t ($cy - $h/2 - 20) -w 120 -h 16 -text "$($n.label)" -size 8.5 -align 2 | Out-Null
                    } else {
                        New-Label -l ($x - 132) -t ($cy + $h/2 + 3) -w 120 -h 14 -text "$($n.label)" -size 8.5 -align 3 | Out-Null
                    }
                }
            }
            "start" {
                $d = 26
                $shp = New-Box $msoShapeOval ($x - $d/2) ($cy - $d/2) $d $d
                Set-BlackLine $shp 1.25
            }
            "end" {
                $d = 26
                $shp = New-Box $msoShapeOval ($x - $d/2) ($cy - $d/2) $d $d
                Set-BlackLine $shp 2.75
            }
            "datastore" {
                # ラベルが折り返して不格好にならないよう、文字数に応じて幅を広げる
                $w = if ($n.width) { [double]$n.width } else { [math]::Max(62, [math]::Min(110, "$($n.label)".Length * 9 + 12)) }
                $h = 50
                $shp = New-Box $msoShapeCan ($x - $w/2) ($cy - $h/2) $w $h
                Set-ShapeText $shp "$($n.label)" 8.5 2
            }
            { $_ -in "dataobject", "document" } {
                # BPMNデータオブジェクト: 角折れ四角の小アイコン＋右横ラベル
                # 書類は大きな図形にせず、線でつながず流れの近くに置く
                $iw = 26; $ih = 32
                $shp = New-Box $msoShapeFoldedCorner ($x - $iw/2) ($cy - $ih/2) $iw $ih
                # ラベルはアイコンの右横・縦中央合わせ（\n 改行対応）
                $lblText = "$($n.label)"
                $lblLines = ($lblText -split "\r?\n")
                $maxLen = 0
                foreach ($ll in $lblLines) { if ($ll.Length -gt $maxLen) { $maxLen = $ll.Length } }
                $lw = $maxLen * 9 + 10
                $lh = $lblLines.Count * 14 + 4
                $lbl = New-Label -l ($x + 13 + 4) -t ($cy - $lh/2) -w $lw -h $lh -text $lblText -size 8.5 -align 1
                $lbl.TextFrame2.WordWrap = 0
            }
            "annotation" {
                $w = if ($n.width) { [double]$n.width } else { 130 }
                $lines = ("$($n.label)" -split "\r?\n").Count
                $h = 17 * $lines + 8
                $shp = $ws.Shapes.AddTextbox(1, $x - $w/2 + 8, $cy - $h/2, $w, $h)
                $shp.Fill.Visible = 0
                $shp.Line.Visible = 0
                Set-ShapeText $shp "$($n.label)" 8.5 1
                # 幅だけ実テキストに合わせ、高さは行数から明示確保する
                # （Excel の自動サイズは日本語フォントの行高を過小評価し、PDF出力時に下の行が切れる）
                $shp.TextFrame2.WordWrap = 0
                $shp.TextFrame2.AutoSize = 1
                $shp.TextFrame2.AutoSize = 0
                $shp.Width = $shp.Width + 6
                $shp.Height = $h
                # 左に BPMN 注記のブラケット
                $br = $ws.Shapes.AddShape($msoShapeLeftBracket, $shp.Left - 7, $shp.Top, 6, $shp.Height)
                $br.Fill.Visible = 0
                Set-BlackLine $br 1.0
                $br.Shadow.Visible = 0
            }
            default { throw "未知のノード種別: $($n.type)（id=$($n.id)）" }
        }
        $nodeShapes[$n.id] = @{ Shape = $shp; X = $x; Y = $cy; Lane = $n.lane; Type = "$($n.type)" }
    }

    # ---- フロー（コネクタ） ----
    foreach ($fl in $def.flows) {
        if (-not $nodeShapes.ContainsKey($fl.from)) { throw "flow の from '$($fl.from)' が nodes に存在しません" }
        if (-not $nodeShapes.ContainsKey($fl.to))   { throw "flow の to '$($fl.to)' が nodes に存在しません" }
        $src = $nodeShapes[$fl.from]; $dst = $nodeShapes[$fl.to]
        $isAssoc = ("$($fl.type)" -eq "association")
        $dx = $dst.X - $src.X; $dy = $dst.Y - $src.Y

        $ctype = if ([math]::Abs($dx) -gt 6 -and [math]::Abs($dy) -gt 6) { $msoConnectorElbow } else { $msoConnectorStraight }
        $conn = $ws.Shapes.AddConnector($ctype, $src.X, $src.Y, $dst.X, $dst.Y)

        $manualSites = ($fl.fromSite -or $fl.toSite)
        $fromSite = if ($fl.fromSite) { $SiteMap["$($fl.fromSite)"] } else { 1 }
        $toSite   = if ($fl.toSite)   { $SiteMap["$($fl.toSite)"] }   else { 1 }
        if (-not $manualSites -and -not $isAssoc -and $src.Lane -ne $dst.Lane -and
            $src.Type -ne "gateway" -and
            [math]::Abs($dx) -gt 6 -and [math]::Abs($dy) -gt 6) {
            # レーンをまたぐ順序フローは横から出て横に入れる
            # （最短ルート任せだと、真下に置いた書類・システムを縦線が貫通する）
            # ゲートウェイ発の分岐は例外: 下方向の枝は下辺から出る方が読みやすい
            if ($dx -gt 0) { $fromSite = 4; $toSite = 2 } else { $fromSite = 2; $toSite = 4 }
            $manualSites = $true
        }
        $conn.ConnectorFormat.BeginConnect($src.Shape, $fromSite)
        $conn.ConnectorFormat.EndConnect($dst.Shape, $toSite)
        if (-not $manualSites) { $conn.RerouteConnections() }

        Set-BlackLine $conn 1.0
        $conn.Shadow.Visible = 0
        if ($isAssoc) {
            $conn.Line.DashStyle = $msoLineDash
        } else {
            $conn.Line.EndArrowheadStyle = $msoArrowheadTriangle
        }

        if ($fl.label) {
            # ラベルは線の中間ではなく出口の近くに置く（BPMNの慣例。行き先ノードとの重なりも防ぐ）
            $lw = "$($fl.label)".Length * 9 + 12
            if ($src.Type -eq "gateway" -and $dy -gt 20) {
                # 下方向の分岐: 縦線のすぐ右
                $lx = $src.X + 8; $ly = $src.Y + 22
            } elseif ($src.Type -eq "gateway" -and $dy -lt -20) {
                # 上方向の分岐: 縦線のすぐ右・上
                $lx = $src.X + 8; $ly = $src.Y - 34
            } elseif ([math]::Abs($dx) -ge 6) {
                # 横方向: 隣接ノードの番号帯と重ならないよう1段上げる
                $lx = if ($dx -gt 0) { $src.X + 24 } else { $src.X - 24 - $lw }
                $ly = $src.Y - 32
            } else {
                $lx = $src.X + 8
                $ly = $src.Y + $dy * 0.3 - 7
            }
            $lbl = New-Label -l $lx -t $ly -w $lw -h 14 -text "$($fl.label)" -size 8.5 -align 1
            $lbl.TextFrame2.WordWrap = 0
        }
    }

    # ---- 脚注（機能要件対応表など） ----
    if ($def.footer) {
        $fw = 260
        $foot = $ws.Shapes.AddTextbox(1, $DiagRight - $fw, $diagBottom + 8, $fw, 40)
        $foot.Fill.ForeColor.RGB = 16777215; $foot.Fill.Visible = -1
        Set-BlackLine $foot 1.0
        $foot.Line.DashStyle = $msoLineDash
        Set-ShapeText $foot "$($def.footer)" 8.5 1
        $flines = ("$($def.footer)" -split "\r?\n").Count
        $foot.Height = 17 * $flines + 8
    }

    # ---- 印刷設定（PDF検証用に図全体を1ページへ） ----
    $lastCol = [int][math]::Ceiling($DiagRight / 48) + 1
    $lastRow = [int][math]::Ceiling(($diagBottom + 70) / 15) + 1
    $colLetter = ""
    $c = $lastCol
    while ($c -gt 0) { $rem = [int](($c - 1) % 26); $colLetter = [string][char](65 + $rem) + $colLetter; $c = [int][math]::Floor(($c - 1) / 26) }
    $ws.PageSetup.PrintArea = "A1:$colLetter$lastRow"
    $ws.PageSetup.Orientation = 2
    $ws.PageSetup.Zoom = $false
    $ws.PageSetup.FitToPagesWide = 1
    $ws.PageSetup.FitToPagesTall = 1

    # ---- 保存 ----
    $wb.SaveAs($outPath, 51)
    Write-Output "生成完了: $outPath"

    if ($ExportPdf) {
        $pdfPath = [System.IO.Path]::ChangeExtension($outPath, ".pdf")
        $wb.ExportAsFixedFormat(0, $pdfPath)
        Write-Output "PDF出力: $pdfPath"
    }

    $wb.Close($false)
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect()
}
