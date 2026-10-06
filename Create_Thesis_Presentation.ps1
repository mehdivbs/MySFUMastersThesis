$ErrorActionPreference = 'Stop'

$ws = 'D:\Academics\SFU Thesis\MySFUMastersThesis'
$output = Join-Path $ws 'Human-Centric_Data-Driven_Decision_Support_20Slides.pptx'

if (Test-Path $output) {
    Remove-Item $output -Force
}

$ppApp = New-Object -ComObject PowerPoint.Application
$ppApp.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
$presentation = $ppApp.Presentations.Add()
$presentation.PageSetup.SlideWidth = 12192000
$presentation.PageSetup.SlideHeight = 6858000

$theme = $presentation.Designs.Item(1)
$fontTitle = 'Aptos'
$fontBody = 'Aptos'
$blue = 0x005B82
$navy = 0x0D1F3E
$teal = 0x009C94
$orange = 0xF17A35
$gray = 0x58606C
$light = 0xF5F7F9
$white = 0xFFFFFF

function Add-Slide([string]$title, [string]$subtitle = '', [bool]$footer = $true) {
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
    $slide.Background.Fill.ForeColor.RGB = $white

    $bar = $slide.Shapes.AddShape(1, 0, 0, 12192000, 90)
    $bar.Fill.ForeColor.RGB = $blue
    $bar.Line.ForeColor.RGB = $blue

    $top = $slide.Shapes.AddTextbox(1, 350000, 180000, 11100000, 500000)
    $top.TextFrame.TextRange.Text = $title
    $top.TextFrame.TextRange.Font.Name = $fontTitle
    $top.TextFrame.TextRange.Font.Size = 26
    $top.TextFrame.TextRange.Font.Bold = $true
    $top.TextFrame.TextRange.Font.Color.RGB = $navy
    $top.TextFrame.Paragraphs(1).Alignment = 1

    if ($subtitle) {
        $sub = $slide.Shapes.AddTextbox(1, 350000, 740000, 11100000, 270000)
        $sub.TextFrame.TextRange.Text = $subtitle
        $sub.TextFrame.TextRange.Font.Name = $fontBody
        $sub.TextFrame.TextRange.Font.Size = 11
        $sub.TextFrame.TextRange.Font.Color.RGB = $gray
        $sub.TextFrame.Paragraphs(1).Alignment = 1
    }

    if ($footer) {
        $f = $slide.Shapes.AddTextbox(1, 350000, 6620000, 11400000, 220000)
        $f.TextFrame.TextRange.Text = 'Human-Centric Data-Driven Decision Support | SFU Master’s Thesis'
        $f.TextFrame.TextRange.Font.Name = $fontBody
        $f.TextFrame.TextRange.Font.Size = 8
        $f.TextFrame.TextRange.Font.Color.RGB = $gray
        $f.TextFrame.Paragraphs(1).Alignment = 2
    }

    return $slide
}

function Add-TextBox($slide, [int]$left, [int]$top, [int]$width, [int]$height, [string]$text, [int]$fontSize = 18, [int]$color = 0, [int]$bold = 0, [int]$align = 0) {
    $tb = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
    $tb.TextFrame.TextRange.Text = $text
    $tb.TextFrame.TextRange.Font.Name = $fontBody
    $tb.TextFrame.TextRange.Font.Size = $fontSize
    $tb.TextFrame.TextRange.Font.Bold = $bold
    $tb.TextFrame.TextRange.Font.Color.RGB = if ($color -eq 0) { $navy } else { $color }
    $tb.TextFrame.Paragraphs(1).Alignment = $align
    return $tb
}

function Add-Image($slide, [string]$path, [int]$left, [int]$top, [int]$width, [int]$height) {
    $img = $slide.Shapes.AddPicture($path, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, $left, $top, $width, $height)
    $img.LockAspectRatio = [Microsoft.Office.Core.MsoTriState]::msoTrue
    return $img
}

function Add-BulletBox($slide, [int]$left, [int]$top, [int]$width, [int]$height, [string[]]$items, [string]$title = '') {
    $box = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    $box.Fill.ForeColor.RGB = $light
    $box.Line.ForeColor.RGB = $light
    $box.TextFrame.TextRange.Text = ''
    if ($title) {
        $r = $box.TextFrame.TextRange
        $r.Text = $title
        $r.Font.Name = $fontTitle
        $r.Font.Size = 18
        $r.Font.Bold = $true
        $r.Font.Color.RGB = $blue
        $r.Paragraphs(1).Alignment = 1
        $r.InsertAfter("`n")
    }
    $p = $box.TextFrame.Paragraphs.Add()
    $p.Text = ''
    $p.Level = 0
    $body = $box.TextFrame.TextRange
    $body.Font.Name = $fontBody
    $body.Font.Size = 13
    $body.Font.Color.RGB = $navy
    foreach ($item in $items) {
        $p = $box.TextFrame.Paragraphs.Add()
        $p.Text = '• ' + $item
        $p.Level = 0
        $p.Bullet = $true
    }
    return $box
}

function Add-Chart($slide, [int]$left, [int]$top, [int]$width, [int]$height, [string]$chartType, [string[]]$labels, [double[]]$values, [string]$titleText) {
    $chart = $slide.Shapes.AddChart($left, $top, $width, $height).Chart
    $chart.ChartType = $chartType
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $titleText
    $chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
    $chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = $true
    $chart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = $fontBody
    $chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $navy

    if ($chartType -eq 3) {
        $chart.SeriesCollection(1).Name = 'Share'
        $chart.SeriesCollection(1).Values = $values
        $chart.SeriesCollection(1).XValues = $labels
        $chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = $teal
    } else {
        $chart.SeriesCollection(1).Name = 'Value'
        $chart.SeriesCollection(1).Values = $values
        $chart.SeriesCollection(1).XValues = $labels
        $chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = $teal
    }

    foreach ($s in $chart.SeriesCollection()) {
        $s.Format.Fill.ForeColor.RGB = $blue
    }
    return $chart
}

$imgDir = Join-Path $ws 'Figures'

# Slide 1
$slide = Add-Slide 'Human-Centric Design for Data-Driven Decision Support' 'A 20-slide presentation of the SFU Master’s thesis' $false
Add-TextBox -slide $slide -left 350000 -top 1450000 -width 11200000 -height 700000 -text 'Mehdi Vahabisani`nMaster of Science Thesis' -fontSize 22 -color $navy -bold 1 -align 1
$tag = $slide.Shapes.AddShape(1, 500000, 2500000, 11200000, 500000)
$tag.Fill.ForeColor.RGB = $light
$tag.Line.ForeColor.RGB = $light
$tag.TextFrame.TextRange.Text = 'Decision intelligence • Causal reasoning • Human-centered design'
$tag.TextFrame.TextRange.Font.Name = $fontBody
$tag.TextFrame.TextRange.Font.Size = 15
$tag.TextFrame.TextRange.Font.Color.RGB = $blue
$tag.TextFrame.Paragraphs(1).Alignment = 1
$slide.Shapes.AddShape(1, 3550000, 3920000, 5100000, 200000).Fill.ForeColor.RGB = $teal
$slide.Shapes.AddShape(1, 3550000, 3920000, 5100000, 200000).Line.ForeColor.RGB = $teal
$slide.Shapes.AddShape(1, 3600000, 3930000, 5000000, 180000).TextFrame.TextRange.Text = 'Data + human judgment + causal understanding'
$slide.Shapes.AddShape(1, 3600000, 3930000, 5000000, 180000).TextFrame.TextRange.Font.Size = 12
$slide.Shapes.AddShape(1, 3600000, 3930000, 5000000, 180000).TextFrame.TextRange.Font.Color.RGB = $white
$slide.Shapes.AddShape(1, 3600000, 3930000, 5000000, 180000).TextFrame.Paragraphs(1).Alignment = 1

# Slide 2
$slide = Add-Slide 'Thesis at a Glance' 'From fragmented tools to an integrated decision-support framework'
Add-BulletBox -slide $slide -left 500000 -top 1200000 -width 4100000 -height 4200000 -items @('What decision-makers need','How data-driven tools help','How causal reasoning and human factors fit','How tools are orchestrated') -title 'Core argument'
Add-BulletBox -slide $slide -left 4900000 -top 1200000 -width 7000000 -height 4200000 -items @('Decision support must be task-aware','No single technique is sufficient','Human agency remains essential','Causal, visual, predictive, and optimization tools complement each other') -title 'Research contribution'

# Slide 3
$slide = Add-Slide 'The Problem' 'Decision support is often fragmented and difficult to trust'
Add-BulletBox -slide $slide -left 500000 -top 1200000 -width 5300000 -height 4400000 -items @('Data tools are applied separately','Correlations can be confused with causal effects','High-dimensional outputs create cognitive overload','Black-box systems may reduce trust and confidence') -title 'Core challenges'
Add-Image -slide $slide -path (Join-Path $imgDir 'DDDMProcess.png') -left 6200000 -top 1250000 -width 5100000 -height 4200000

# Slide 4
$slide = Add-Slide 'Research Questions' 'The thesis asks how data-driven tools and human reasoning can work together'
Add-BulletBox -slide $slide -left 500000 -top 1200000 -width 5400000 -height 4400000 -items @('RQ1: Which data-driven techniques support decisions?','RQ2: How do these techniques augment decision-making?','RQ3: What role does causal reasoning play?','RQ4: How should tools be integrated for the user and task?') -title 'Four research questions'
Add-Image -slide $slide -path (Join-Path $imgDir 'ResearchProcess.png') -left 6200000 -top 1250000 -width 5100000 -height 4200000

# Slide 5
$slide = Add-Slide 'Research Methodology' 'Literature review, thematic analysis, and framework synthesis'
Add-BulletBox -slide $slide -left 500000 -top 1200000 -width 5400000 -height 4400000 -items @('Review data-driven decision-making literature','Classify recurring techniques','Analyze causal reasoning and human factors','Synthesize a human-and-task-centered framework') -title 'Methodological approach'
Add-Image -slide $slide -path (Join-Path $imgDir 'ResearchProcess.png') -left 6200000 -top 1250000 -width 5100000 -height 4200000

# Slide 6
$slide = Add-Slide 'The Literature Landscape' 'Six major classes of data-driven techniques were identified'
Add-BulletBox -slide $slide -left 500000 -top 1180000 -width 2400000 -height 4300000 -items @('Machine learning','Traditional statistics','Data visualization','Natural language models','Mathematical optimization','Simulation') -title 'Technique classes'
Add-Image -slide $slide -path (Join-Path $imgDir 'DDDMLitReview_search_summary.png') -left 3200000 -top 1180000 -width 8200000 -height 4300000

# Slide 7
$slide = Add-Slide 'The Decision-Making Process' 'Simon’s three-stage model provides the organizing structure'
Add-BulletBox -slide $slide -left 500000 -top 1300000 -width 3300000 -height 3600000 -items @('Intelligence: detect and frame the problem','Design: define criteria, variables, and alternatives','Choice: evaluate, compare, and select') -title 'Three stages'
Add-Image -slide $slide -path (Join-Path $imgDir 'DDTAssignment.png') -left 4200000 -top 1300000 -width 7600000 -height 3800000

# Slide 8
$slide = Add-Slide 'Data-Driven Techniques' 'Machine learning supports prediction, grouping, anomalies, and adaptation'
Add-BulletBox -slide $slide -left 450000 -top 1240000 -width 3600000 -height 4150000 -items @('Supervised learning: prediction and classification','Unsupervised learning: clustering and anomaly detection','Reinforcement learning: adaptive policies','Explainable models: transparent recommendations') -title 'Machine learning'
Add-Image -slide $slide -path (Join-Path $imgDir 'MCDMTree.jpg') -left 4350000 -top 1300000 -width 7100000 -height 3700000

# Slide 9
$slide = Add-Slide 'Traditional Statistics' 'Statistics provide interpretable baselines and uncertainty estimates'
Add-BulletBox -slide $slide -left 500000 -top 1200000 -width 4400000 -height 4300000 -items @('Descriptive and exploratory analysis','Hypothesis testing','Forecasting and time-series analysis','Risk estimation and threshold monitoring') -title 'Strengths'
Add-Image -slide $slide -path (Join-Path $imgDir 'ContingencyTables.png') -left 5400000 -top 1250000 -width 6100000 -height 3500000

# Slide 10
$slide = Add-Slide 'Data Visualization' 'Visual analytics turns complex data into interactive evidence'
Add-BulletBox -slide $slide -left 500000 -top 1200000 -width 3800000 -height 4100000 -items @('Pattern recognition','Comparison of alternatives','Temporal and spatial exploration','Interactive cross-filtering and exploration') -title 'Why visualization matters'
Add-Image -slide $slide -path (Join-Path $imgDir 'VADashboardPublicHealth.png') -left 4700000 -top 1200000 -width 6800000 -height 4100000

# Slide 11
$slide = Add-Slide 'Optimization and Simulation' 'Formal methods evaluate trade-offs and hypothetical futures'
Add-BulletBox -slide $slide -left 500000 -top 1200000 -width 3500000 -height 4300000 -items @('Multi-objective optimization: Pareto trade-offs','MCDM: rank alternatives across criteria','Simulation: scenario and sensitivity analysis','Stochastic models: represent uncertainty') -title 'Decision models'
Add-Image -slide $slide -path (Join-Path $imgDir 'Biobjective_VIZ.png') -left 4400000 -top 1200000 -width 7000000 -height 4300000

# Slide 12
$slide = Add-Slide 'Causal Reasoning' 'Understanding why an outcome occurs is essential for intervention'
Add-BulletBox -slide $slide -left 500000 -top 1250000 -width 3500000 -height 3900000 -items @('Causes, effects, mediators, and confounders','Action-oriented reasoning','Counterfactual (“what if?”) thinking','Causal models support intervention planning') -title 'Human reasoning'
Add-Image -slide $slide -path (Join-Path $imgDir 'CausalGraphConstructionProcess.png') -left 4400000 -top 1250000 -width 6900000 -height 3900000

# Slide 13
$slide = Add-Slide 'Counterfactual and What-If Thinking' 'Decision-makers compare actual outcomes with plausible alternatives'
Add-BulletBox -slide $slide -left 500000 -top 1220000 -width 3300000 -height 3900000 -items @('What would have happened without this action?','Which action would have produced a better outcome?','How robust is the selected option?','How does uncertainty change the comparison?') -title 'Counterfactual questions'
Add-Image -slide $slide -path (Join-Path $imgDir 'DAGExample.png') -left 4300000 -top 1200000 -width 7100000 -height 4000000

# Slide 14
$slide = Add-Slide 'Humanistic Factors' 'Decision quality depends on expertise, knowledge, and cognition'
Add-BulletBox -slide $slide -left 500000 -top 1230000 -width 3300000 -height 4200000 -items @('Working memory and attention','Domain knowledge and mental models','Expertise and experience','Decision style and cognitive preferences') -title 'Individual differences'
Add-BulletBox -slide $slide -left 4300000 -top 1230000 -width 7200000 -height 4200000 -items @('Analytical users benefit from structured evidence','Intuitive users may prefer fast, experience-based guidance','Adaptive experts can revise assumptions','A single interface should not enforce one decision style') -title 'Design implications'

# Slide 15
$slide = Add-Slide 'Human-and-Task-Centered Framework' 'The framework aligns tools with decision stage, task, and user'
Add-Image -slide $slide -path (Join-Path $imgDir 'ThreeLayerIntegration.png') -left 500000 -top 1200000 -width 11100000 -height 4700000

# Slide 16
$slide = Add-Slide 'Intelligence Stage' 'Discover signals, frame the problem, and reduce information overload'
Add-Image -slide $slide -path (Join-Path $imgDir 'IntelligenceWorkflow.png') -left 300000 -top 1300000 -width 11500000 -height 4500000

# Slide 17
$slide = Add-Slide 'Design Stage' 'Build criteria, generate alternatives, and model causal relationships'
Add-Image -slide $slide -path (Join-Path $imgDir 'DesignWorkflow_ArrowsAmplified.png') -left 300000 -top 1300000 -width 11500000 -height 4500000

# Slide 18
$slide = Add-Slide 'Choice Stage' 'Compare outcomes, explore trade-offs, and select a robust action'
Add-Image -slide $slide -path (Join-Path $imgDir 'ChoiceWorkflow.png') -left 300000 -top 1300000 -width 11500000 -height 4500000

# Slide 19
$slide = Add-Slide 'ICADOS Architecture' 'A conceptual decision-support system integrating the techniques'
Add-BulletBox -slide $slide -left 500000 -top 1250000 -width 4200000 -height 4200000 -items @('Causal and task specification','Evidence and model generation','Interactive decision support','Outcome simulation and comparison') -title 'System components'
Add-Image -slide $slide -path (Join-Path $imgDir 'ICADOS_UCD.png') -left 5100000 -top 1250000 -width 6400000 -height 4200000

# Slide 20
$slide = Add-Slide 'Key Takeaways' 'A practical summary for future decision-support design' $false
Add-BulletBox -slide $slide -left 500000 -top 1200000 -width 4700000 -height 4100000 -items @('Use tools for the decision task, not as a generic solution','Keep causal structure visible and testable','Support user reasoning rather than replace it','Combine prediction, optimization, simulation, and visualization','Evaluate recommendations for robustness, trust, and usability') -title 'Takeaways'
$chart = $slide.Shapes.AddChart(5400000, 1350000, 5400000, 3700000).Chart
$chart.ChartType = 4
$chart.HasTitle = $true
$chart.ChartTitle.Text = 'Illustrative tool emphasis'
$chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
$chart.ChartTitle.Format.TextFrame2.TextRange.Font.Name = $fontBody
$chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = $true
$chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = $navy
$chart.SeriesCollection(1).Name = 'Importance'
$chart.SeriesCollection(1).Values = @(1, 3, 4, 5, 4)
$chart.SeriesCollection(1).XValues = @('ML', 'TS', 'VIZ', 'MO', 'SIM')
$chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = $teal
$chart.SeriesCollection(1).Format.Line.ForeColor.RGB = $teal
$chart.ApplyDataLabels(8)

$slide.Shapes.AddTextbox(1, 5400000, 5300000, 5400000, 600000).TextFrame.TextRange.Text = 'A complete deck can be extended with the author’s own case study data.'
$slide.Shapes.AddTextbox(1, 5400000, 5300000, 5400000, 600000).TextFrame.TextRange.Font.Name = $fontBody
$slide.Shapes.AddTextbox(1, 5400000, 5300000, 5400000, 600000).TextFrame.TextRange.Font.Size = 10
$slide.Shapes.AddTextbox(1, 5400000, 5300000, 5400000, 600000).TextFrame.TextRange.Font.Color.RGB = $gray
$slide.Shapes.AddTextbox(1, 5400000, 5300000, 5400000, 600000).TextFrame.Paragraphs(1).Alignment = 2

$presentation.SaveAs($output)
$presentation.Close()
$ppApp.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($chart) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppApp) | Out-Null

Write-Output "Created $output"
