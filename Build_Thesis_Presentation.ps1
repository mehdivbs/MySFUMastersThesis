$ErrorActionPreference = 'Stop'

$root = 'D:\Academics\SFU Thesis\MySFUMastersThesis'
$out = Join-Path $root 'Human-Centric_Data-Driven_Decision_Support_20Slides.pptx'
$figures = Join-Path $root 'Figures'

if (Test-Path $out) { Remove-Item $out -Force }

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.DisplayAlerts = [Microsoft.Office.Interop.PowerPoint.PpAlertLevel]::ppAlertsNone
$deck = $ppt.Presentations.Add()
$deck.PageSetup.SlideWidth = 12000
$deck.PageSetup.SlideHeight = 6750

function New-Slide($title, $subtitle = '', $showFooter = $true) {
    $slide = $deck.Slides.Add($deck.Slides.Count + 1, 12)
    $slide.Background.Fill.ForeColor.RGB = 0xFFFFFF

    $bar = $slide.Shapes.AddShape(1, 0, 0, 12000, 120)
    $bar.Fill.ForeColor.RGB = 0x005B82
    $bar.Line.ForeColor.RGB = 0x005B82

    $tb = $slide.Shapes.AddTextbox(1, 500, 180, 11000, 460)
    $tb.TextFrame.TextRange.Text = $title
    $tb.TextFrame.TextRange.Font.Name = 'Aptos'
    $tb.TextFrame.TextRange.Font.Size = 26
    $tb.TextFrame.TextRange.Font.Bold = $true
    $tb.TextFrame.TextRange.Font.Color.RGB = 0x0D1F3E
    $tb.TextFrame.Paragraphs(1).Alignment = 2

    if ($subtitle) {
        $tb2 = $slide.Shapes.AddTextbox(1, 500, 660, 11000, 250)
        $tb2.TextFrame.TextRange.Text = $subtitle
        $tb2.TextFrame.TextRange.Font.Name = 'Aptos'
        $tb2.TextFrame.TextRange.Font.Size = 10
        $tb2.TextFrame.TextRange.Font.Color.RGB = 0x58606C
        $tb2.TextFrame.Paragraphs(1).Alignment = 2
    }

    if ($showFooter) {
        $f = $slide.Shapes.AddTextbox(1, 500, 6300, 11000, 180)
        $f.TextFrame.TextRange.Text = 'Human-Centric Data-Driven Decision Support | SFU Master’s Thesis'
        $f.TextFrame.TextRange.Font.Name = 'Aptos'
        $f.TextFrame.TextRange.Font.Size = 8
        $f.TextFrame.TextRange.Font.Color.RGB = 0x58606C
        $f.TextFrame.Paragraphs(1).Alignment = 2
    }

    return $slide
}

function Add-Text($slide, $x, $y, $w, $h, $text, $size = 18, $color = 0x0D1F3E, $bold = $false, $align = 0) {
    $tb = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $tb.TextFrame.TextRange.Text = $text
    $tb.TextFrame.TextRange.Font.Name = 'Aptos'
    $tb.TextFrame.TextRange.Font.Size = $size
    $tb.TextFrame.TextRange.Font.Bold = $bold
    $tb.TextFrame.TextRange.Font.Color.RGB = $color
    $tb.TextFrame.Paragraphs(1).Alignment = $align
    return $tb
}

function Add-Bullets($slide, $x, $y, $w, $h, $items, $title = '') {
    $box = $slide.Shapes.AddShape(1, $x, $y, $w, $h)
    $box.Fill.ForeColor.RGB = 0xF5F7F9
    $box.Line.ForeColor.RGB = 0xF5F7F9
    $box.TextFrame.TextRange.Text = $title
    $box.TextFrame.TextRange.Font.Name = 'Aptos'
    $box.TextFrame.TextRange.Font.Size = 18
    $box.TextFrame.TextRange.Font.Bold = $true
    $box.TextFrame.TextRange.Font.Color.RGB = 0x005B82
    $box.TextFrame.Paragraphs(1).Alignment = 2
    foreach ($item in $items) {
        $p = $box.TextFrame.Paragraphs.Add()
        $p.Text = '• ' + $item
        $p.Level = 0
    }
    $box.TextFrame.TextRange.Font.Name = 'Aptos'
    $box.TextFrame.TextRange.Font.Size = 13
    $box.TextFrame.TextRange.Font.Color.RGB = 0x0D1F3E
    return $box
}

function Add-Image($slide, $path, $x, $y, $w, $h) {
    $img = $slide.Shapes.AddPicture($path, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, $x, $y, $w, $h)
    $img.LockAspectRatio = [Microsoft.Office.Core.MsoTriState]::msoTrue
    return $img
}

function Add-LineChart($slide, $x, $y, $w, $h, $labels, $values, $title) {
    $chart = $slide.Shapes.AddChart($x, $y, $w, $h).Chart
    $chart.ChartType = 4
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $title
    $chart.SeriesCollection(1).Name = 'Importance'
    $chart.SeriesCollection(1).Values = $values
    $chart.SeriesCollection(1).XValues = $labels
    $chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = 0x009C94
    $chart.SeriesCollection(1).Format.Line.ForeColor.RGB = 0x009C94
    return $chart
}

$slides = @(
    @{ title='Human-Centric Design for Data-Driven Decision Support'; subtitle='A 20-slide presentation of the SFU Master’s thesis'; footer=$false },
    @{ title='Thesis at a Glance'; subtitle='From fragmented tools to an integrated decision-support framework' },
    @{ title='The Research Problem'; subtitle='Decision support is fragmented and hard to trust' },
    @{ title='Research Questions'; subtitle='Four questions guide the thesis' },
    @{ title='Research Methodology'; subtitle='Literature review, thematic analysis, and synthesis' },
    @{ title='The Literature Landscape'; subtitle='Six major classes of data-driven techniques' },
    @{ title='The Decision-Making Process'; subtitle='Simon’s three-stage model' },
    @{ title='Machine Learning'; subtitle='Prediction, classification, grouping, and adaptation' },
    @{ title='Traditional Statistics'; subtitle='Interpretability, hypotheses, forecasting, and risk' },
    @{ title='Data Visualization'; subtitle='Interactive visual analytics for sense-making' },
    @{ title='Optimization and Simulation'; subtitle='Trade-offs, scenarios, and uncertainty' },
    @{ title='Causal Reasoning'; subtitle='From correlation to intervention and explanation' },
    @{ title='Counterfactual Thinking'; subtitle='Comparing real outcomes with hypothetical alternatives' },
    @{ title='Humanistic Factors'; subtitle='Cognition, knowledge, expertise, and decision style' },
    @{ title='Human-and-Task-Centered Framework'; subtitle='Aligning tools with the decision stage and user' },
    @{ title='Intelligence Stage'; subtitle='Detecting signals and framing the problem' },
    @{ title='Design Stage'; subtitle='Defining criteria, alternatives, and causal structures' },
    @{ title='Choice Stage'; subtitle='Evaluating outcomes, trade-offs, and sensitivity' },
    @{ title='ICADOS Architecture'; subtitle='A conceptual system integrating data-driven tools and causal reasoning' },
    @{ title='Key Takeaways'; subtitle='A practical agenda for future decision-support design'; footer=$false }
)

$slide = New-Slide $slides[0].title $slides[0].subtitle $slides[0].footer
Add-Text $slide 600 1400 11000 350 'Mehdi Vahabisani' 30 $color=0x0D1F3E $bold=$true $align=2
Add-Text $slide 600 1800 11000 250 'Master of Science Thesis' 18 $color=0x58606C $bold=$false $align=2
Add-Text $slide 1800 2850 8200 550 'Data + human judgment + causal understanding' 18 $color=0xFFFFFF $bold=$true $align=1
$tag = $slide.Shapes.AddShape(1, 1600, 2720, 8800, 540)
$tag.Fill.ForeColor.RGB = 0x009C94
$tag.Line.ForeColor.RGB = 0x009C94
Add-Text $slide 600 4300 11000 240 'Decision intelligence • Causal reasoning • Human-centered design' 14 $color=0x005B82 $bold=$true $align=2

for ($i = 1; $i -lt $slides.Count; $i++) {
    $slide = New-Slide $slides[$i].title $slides[$i].subtitle $slides[$i].footer
    switch ($i) {
        1 { Add-Bullets $slide 500 1100 4400 3900 @('Decision support must be task-aware','No single technique is sufficient','Human agency remains essential','Causal, predictive, and optimization tools complement each other') 'Core argument'; Add-Bullets $slide 5300 1100 6100 3900 @('The framework integrates methods across stages','The thesis is both conceptual and prescriptive','The system should support reasoning, not replace it','The design is intended for organizational decisions') 'Research contribution' }
        2 { Add-Bullets $slide 500 1100 4700 4100 @('Data tools are applied separately','Correlations can be mistaken for causes','High-dimensional outputs increase cognitive load','Black-box models may reduce trust') 'Core challenges'; Add-Image $slide (Join-Path $figures 'DDDMProcess.png') 5700 1200 5600 3800 }
        3 { Add-Bullets $slide 500 1100 4700 4000 @('RQ1: Which tools support decisions?','RQ2: How do they augment decision-making?','RQ3: What role does causal reasoning play?','RQ4: How should tools be integrated for the user and task?') 'Four research questions'; Add-Image $slide (Join-Path $figures 'ResearchProcess.png') 5700 1200 5600 3800 }
        4 { Add-Bullets $slide 500 1100 4700 4000 @('Review data-driven decision-making literature','Classify recurring techniques','Analyze causal reasoning and human factors','Synthesize a human-and-task-centered framework') 'Methodological approach'; Add-Image $slide (Join-Path $figures 'ResearchProcess.png') 5700 1200 5600 3800 }
        5 { Add-Bullets $slide 500 1100 2700 4100 @('Machine learning','Traditional statistics','Data visualization','Natural language models','Mathematical optimization','Simulation') 'Technique classes'; Add-Image $slide (Join-Path $figures 'DDDMLitReview_search_summary.png') 3500 1200 7600 3700 }
        6 { Add-Bullets $slide 500 1100 3300 4000 @('Intelligence: detect and frame the problem','Design: define criteria, variables, and alternatives','Choice: evaluate, compare, and select') 'Three stages'; Add-Image $slide (Join-Path $figures 'DDTAssignment.png') 4200 1200 7000 3800 }
        7 { Add-Bullets $slide 500 1100 3500 4100 @('Supervised learning: prediction & classification','Unsupervised learning: clustering & anomalies','Reinforcement learning: adaptive policies','Explainable models: transparent recommendations') 'Machine learning'; Add-Image $slide (Join-Path $figures 'MCDMTree.jpg') 4350 1200 6800 3600 }
        8 { Add-Bullets $slide 500 1100 3600 4200 @('Descriptive exploration','Hypothesis testing','Time-series forecasting','Risk estimation and threshold monitoring') 'Traditional statistics'; Add-Image $slide (Join-Path $figures 'ContingencyTables.png') 4500 1150 6200 3500 }
        9 { Add-Bullets $slide 500 1100 3600 4200 @('Pattern recognition','Alternative comparison','Temporal and spatial exploration','Interactive cross-filtering') 'Data visualization'; Add-Image $slide (Join-Path $figures 'VADashboardPublicHealth.png') 4500 1150 6200 3500 }
        10 { Add-Bullets $slide 500 1100 3500 4200 @('Multi-objective optimization: Pareto trade-offs','MCDM: ranking across criteria','Simulation: scenario analysis','Stochastic models: represent uncertainty') 'Optimization and simulation'; Add-Image $slide (Join-Path $figures 'Biobjective_VIZ.png') 4400 1150 6500 3700 }
        11 { Add-Bullets $slide 500 1100 3500 4200 @('Causes, effects, mediators, and confounders','Action-oriented reasoning','Counterfactual what-if thinking','Causal models support interventions') 'Causal reasoning'; Add-Image $slide (Join-Path $figures 'CausalGraphConstructionProcess.png') 4300 1200 6600 3600 }
        12 { Add-Bullets $slide 500 1100 3500 4200 @('What would have happened without this action?','Which action would have produced a better outcome?','How robust is the selected option?','How does uncertainty affect the comparison?') 'Counterfactual questions'; Add-Image $slide (Join-Path $figures 'DAGExample.png') 4300 1200 6600 3700 }
        13 { Add-Bullets $slide 500 1100 3600 4200 @('Working memory and attention','Domain knowledge and mental models','Expertise and experience','Decision style and cognitive preferences') 'Humanistic factors'; Add-Bullets $slide 4500 1200 6200 3900 @('Analytical users benefit from structured evidence','Intuitive users prefer fast guidance','Adaptive experts revise assumptions','Support should not enforce one style') 'Design implications' }
        14 { Add-Image $slide (Join-Path $figures 'ThreeLayerIntegration.png') 500 1100 11000 4100 }
        15 { Add-Image $slide (Join-Path $figures 'IntelligenceWorkflow.png') 400 1100 11200 4100 }
        16 { Add-Image $slide (Join-Path $figures 'DesignWorkflow_ArrowsAmplified.png') 400 1100 11200 4100 }
        17 { Add-Image $slide (Join-Path $figures 'ChoiceWorkflow.png') 400 1100 11200 4100 }
        18 { Add-Bullets $slide 500 1100 3500 4200 @('Causal and task specification','Evidence and model generation','Interactive decision support','Outcome simulation and comparison') 'System components'; Add-Image $slide (Join-Path $figures 'ICADOS_UCD.png') 4350 1200 6600 3600 }
        19 { Add-Bullets $slide 500 1100 3900 4200 @('Use tools for the decision task, not as a generic solution','Keep causal structure visible and testable','Support user reasoning rather than replace it','Combine prediction, optimization, simulation, and visualization','Evaluate recommendations for robustness, trust, and usability') 'Takeaways'; Add-LineChart $slide 5300 1500 5600 3200 @('ML','TS','VIZ','MO','SIM') @(1,3,4,5,4) 'Illustrative tool emphasis' }
    }
}

$deck.SaveAs($out)
$deck.Close()
$ppt.Quit()

Write-Output "Created $out"
