<#
.SYNOPSIS
    Creates an example Word template for use with Generate-RampartReport.ps1.

.DESCRIPTION
    Generates a professionally formatted .docx template with all available
    placeholder variables and table markers pre-populated. This serves as a
    starting point that can be customised in Word.

.PARAMETER OutputPath
    Path for the generated template file (default: template.docx in the script directory).

.EXAMPLE
    .\New-RampartTemplate.ps1
    .\New-RampartTemplate.ps1 -OutputPath "C:\Templates\my-template.docx"
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'template.docx')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

# Colour constants
$NavyBlue  = 1848924   # RGB(28, 58, 92)  = 0x1C3A5C
$DarkGrey  = 3355443   # RGB(51, 51, 51)
$MidGrey   = 10066329  # RGB(153, 153, 153)
$White     = 16777215
$LightGrey = 15921906  # RGB(242, 242, 242) = F2F2F2

# Severity colours
$CriticalRed  = 192       # RGB(192, 0, 0)    = 0xC00000 -> Word uses BGR
$HighOrange   = 597743    # RGB(227, 108, 9)
$MediumYellow = 47104     # RGB(230, 184, 0)
$LowBlue      = 4481732   # RGB(68, 114, 196)

# Word enum values
$wdAlignParagraphCenter = 1
$wdAlignParagraphRight  = 2
$wdAlignParagraphLeft   = 0
$wdOrientPortrait       = 0
$wdLineBreak            = 6
$wdPageBreak            = 7
$wdStory                = 6
$wdSeekMainDocument     = 0
$wdSeekPrimaryHeader    = 1
$wdSeekPrimaryFooter    = 2
$wdFormatXMLDocument    = 12
$wdCellAlignVerticalCenter = 1

function Set-HeaderRow {
    param([object]$Row, [string[]]$Headers)
    $Row.Shading.BackgroundPatternColor = $NavyBlue
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $cell = $Row.Cells.Item($i + 1)
        $cell.Range.Text = $Headers[$i]
        $cell.Range.Font.Bold = $true
        $cell.Range.Font.Size = 9
        $cell.Range.Font.Name = 'Calibri'
        $cell.Range.Font.Color = $White
        $cell.VerticalAlignment = $wdCellAlignVerticalCenter
    }
}

function Set-LabelValueRow {
    param([object]$Row, [string]$Label, [string]$Value)
    $Row.Cells.Item(1).Shading.BackgroundPatternColor = $LightGrey
    $Row.Cells.Item(1).Range.Text = $Label
    $Row.Cells.Item(1).Range.Font.Bold = $true
    $Row.Cells.Item(1).Range.Font.Size = 9
    $Row.Cells.Item(1).Range.Font.Name = 'Calibri'
    $Row.Cells.Item(2).Range.Text = $Value
    $Row.Cells.Item(2).Range.Font.Size = 9
    $Row.Cells.Item(2).Range.Font.Name = 'Calibri'
}

function Add-MarkerRow {
    param([object]$Table, [string]$MarkerName, [string[]]$Placeholders)

    # Start marker row: {{#name}} in first cell, rest blank
    $startRow = $Table.Rows.Add()
    $startRow.Cells.Item(1).Range.Text = "{{#$MarkerName}}"
    $startRow.Cells.Item(1).Range.Font.Size = 8
    $startRow.Cells.Item(1).Range.Font.Color = $MidGrey
    for ($i = 2; $i -le $Placeholders.Count; $i++) {
        $startRow.Cells.Item($i).Range.Text = ''
    }

    # Data row with placeholders
    $dataRow = $Table.Rows.Add()
    for ($i = 0; $i -lt $Placeholders.Count; $i++) {
        $cell = $dataRow.Cells.Item($i + 1)
        $cell.Range.Text = "{{ $($Placeholders[$i]) }}"
        $cell.Range.Font.Size = 9
        $cell.Range.Font.Name = 'Calibri'
    }

    # End marker row: {{/name}} in first cell
    $endRow = $Table.Rows.Add()
    $endRow.Cells.Item(1).Range.Text = "{{/$MarkerName}}"
    $endRow.Cells.Item(1).Range.Font.Size = 8
    $endRow.Cells.Item(1).Range.Font.Color = $MidGrey
    for ($i = 2; $i -le $Placeholders.Count; $i++) {
        $endRow.Cells.Item($i).Range.Text = ''
    }
}

# ---------------------------------------------------------------------------
# Create the document
# ---------------------------------------------------------------------------

$word = $null
try {
    Write-Host 'Creating template with Word...'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Add()

    # Page setup (A4)
    $section = $doc.Sections.Item(1)
    $section.PageSetup.PaperSize = 7  # wdPaperA4
    $section.PageSetup.TopMargin    = $word.InchesToPoints(1.0)
    $section.PageSetup.BottomMargin = $word.InchesToPoints(0.75)
    $section.PageSetup.LeftMargin   = $word.InchesToPoints(1.0)
    $section.PageSetup.RightMargin  = $word.InchesToPoints(1.0)

    # Default font
    $doc.Styles.Item('Normal').Font.Name  = 'Calibri'
    $doc.Styles.Item('Normal').Font.Size  = 11
    $doc.Styles.Item('Normal').Font.Color = $DarkGrey

    # Heading styles
    foreach ($level in 1..3) {
        $style = $doc.Styles.Item("Heading $level")
        $style.Font.Name  = 'Calibri'
        $style.Font.Color = $NavyBlue
    }

    # --- Header ---
    $word.ActiveWindow.ActivePane.View.SeekView = $wdSeekPrimaryHeader
    $headerRange = $section.Headers.Item(1).Range
    $headerRange.ParagraphFormat.Alignment = $wdAlignParagraphRight
    $headerRange.Text = '{{ auditor_company }}  |  {{ confidentiality }}'
    $headerRange.Font.Size  = 8
    $headerRange.Font.Color = $MidGrey
    $headerRange.Font.Name  = 'Calibri'

    # --- Footer ---
    $word.ActiveWindow.ActivePane.View.SeekView = $wdSeekPrimaryFooter
    $footerRange = $section.Footers.Item(1).Range
    $footerRange.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    $footerRange.Text = '{{ report_title }}  |  {{ client_name }}  |  {{ report_date }}'
    $footerRange.Font.Size  = 8
    $footerRange.Font.Color = $MidGrey
    $footerRange.Font.Name  = 'Calibri'

    # Back to main document
    $word.ActiveWindow.ActivePane.View.SeekView = $wdSeekMainDocument

    # =========================================================================
    # COVER PAGE
    # =========================================================================
    $range = $doc.Content

    # Spacer
    for ($i = 0; $i -lt 6; $i++) {
        $range.InsertAfter("`r")
    }
    $range.InsertParagraphAfter()

    # Title
    $titlePara = $doc.Content.Paragraphs.Add()
    $titlePara.Range.Text = '{{ report_title }}'
    $titlePara.Range.Font.Size  = 32
    $titlePara.Range.Font.Bold  = $true
    $titlePara.Range.Font.Color = $NavyBlue
    $titlePara.Range.Font.Name  = 'Calibri'
    $titlePara.Alignment = $wdAlignParagraphCenter
    $titlePara.Range.InsertParagraphAfter()

    # Subtitle
    $subPara = $doc.Content.Paragraphs.Add()
    $subPara.Range.Text = 'Prepared for {{ client_name }}'
    $subPara.Range.Font.Size  = 16
    $subPara.Range.Font.Color = $MidGrey
    $subPara.Range.Font.Name  = 'Calibri'
    $subPara.Alignment = $wdAlignParagraphCenter
    $subPara.Range.InsertParagraphAfter()

    # Spacer
    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # Cover metadata table
    $metaRange = $doc.Content.Paragraphs.Add().Range
    $metaTable = $doc.Tables.Add($metaRange, 6, 2)
    $metaTable.Borders.Enable = $true
    $metaItems = @(
        @('Date',            '{{ report_date }}'),
        @('Client Contact',  '{{ client_contact }}'),
        @('Auditor',         '{{ auditor_name }}'),
        @('Company',         '{{ auditor_company }}'),
        @('Confidentiality', '{{ confidentiality }}'),
        @('Document Version','1.0')
    )
    for ($i = 0; $i -lt $metaItems.Count; $i++) {
        Set-LabelValueRow -Row $metaTable.Rows.Item($i + 1) -Label $metaItems[$i][0] -Value $metaItems[$i][1]
    }
    $metaTable.Columns.Item(1).Width = $word.InchesToPoints(2.0)
    $metaTable.Columns.Item(2).Width = $word.InchesToPoints(3.5)

    # Page break
    $doc.Content.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    # =========================================================================
    # TABLE OF CONTENTS
    # =========================================================================
    $tocPara = $doc.Content.Paragraphs.Add()
    $tocPara.Style = 'Heading 1'
    $tocPara.Range.Text = 'Table of Contents'
    $tocPara.Range.InsertParagraphAfter()

    # Insert TOC field
    $tocRange = $doc.Content.Paragraphs.Add().Range
    $doc.TablesOfContents.Add($tocRange, $true, 1, 3)

    $doc.Content.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    # =========================================================================
    # Helper: add a section heading + paragraph
    # =========================================================================
    function Add-Section {
        param([string]$Title, [int]$Level = 1, [string]$Body = '')
        $p = $doc.Content.Paragraphs.Add()
        $p.Style = "Heading $Level"
        $p.Range.Text = $Title
        $p.Range.InsertParagraphAfter()
        if ($Body) {
            $bp = $doc.Content.Paragraphs.Add()
            $bp.Style = 'Normal'
            $bp.Range.Text = $Body
            $bp.Range.InsertParagraphAfter()
        }
    }

    # =========================================================================
    # 1. EXECUTIVE SUMMARY
    # =========================================================================
    Add-Section '1. Executive Summary' -Body (
        'This report presents the findings of a firewall security audit conducted ' +
        'for {{ client_name }} on {{ report_date }}. The audit analysed ' +
        '{{ total_rules }} firewall rules across {{ device_group_count }} device ' +
        'group(s) ({{ device_groups }}).'
    )

    # 1.1 Key Metrics
    Add-Section '1.1 Key Metrics' -Level 2

    $metricsRange = $doc.Content.Paragraphs.Add().Range
    $metricsTable = $doc.Tables.Add($metricsRange, 4, 4)
    $metricsTable.Borders.Enable = $true
    $metrics = @(
        @('Total Rules',        '{{ total_rules }}',          'Compliance Rate',      '{{ compliance_rate }}'),
        @('Rules with Issues',  '{{ rules_with_issues }}',    'Risk Score',           '{{ risk_score }}'),
        @('Total Findings',     '{{ total_findings }}',       'Risk Grade',           '{{ risk_grade }}'),
        @('Shadowed Rules',     '{{ shadowed_rule_count }}',  'Best Practices Score', '{{ best_practices_overall_score }}')
    )
    for ($i = 0; $i -lt 4; $i++) {
        $row = $metricsTable.Rows.Item($i + 1)
        $row.Cells.Item(1).Shading.BackgroundPatternColor = $LightGrey
        Set-LabelValueRow -Row $row -Label $metrics[$i][0] -Value $metrics[$i][1]
        $row.Cells.Item(3).Shading.BackgroundPatternColor = $LightGrey
        $row.Cells.Item(3).Range.Text = $metrics[$i][2]
        $row.Cells.Item(3).Range.Font.Bold = $true
        $row.Cells.Item(3).Range.Font.Size = 9
        $row.Cells.Item(3).Range.Font.Name = 'Calibri'
        $row.Cells.Item(4).Range.Text = $metrics[$i][3]
        $row.Cells.Item(4).Range.Font.Size = 9
        $row.Cells.Item(4).Range.Font.Name = 'Calibri'
    }

    # Spacer
    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # 1.2 Findings by Severity
    Add-Section '1.2 Findings by Severity' -Level 2

    $sevRange = $doc.Content.Paragraphs.Add().Range
    $sevTable = $doc.Tables.Add($sevRange, 1, 2)
    $sevTable.Borders.Enable = $true
    Set-HeaderRow -Row $sevTable.Rows.Item(1) -Headers @('Severity', 'Count')

    $sevItems = @(
        @('Critical', '{{ critical_count }}'),
        @('High',     '{{ high_count }}'),
        @('Medium',   '{{ medium_count }}'),
        @('Low',      '{{ low_count }}'),
        @('Total',    '{{ total_findings }}')
    )
    foreach ($item in $sevItems) {
        $row = $sevTable.Rows.Add()
        $row.Cells.Item(1).Range.Text = $item[0]
        $row.Cells.Item(1).Range.Font.Bold = $true
        $row.Cells.Item(1).Range.Font.Size = 9
        $row.Cells.Item(1).Range.Font.Name = 'Calibri'
        $row.Cells.Item(2).Range.Text = $item[1]
        $row.Cells.Item(2).Range.Font.Size = 9
        $row.Cells.Item(2).Range.Font.Name = 'Calibri'
    }
    $sevTable.Columns.Item(1).Width = $word.InchesToPoints(2.0)
    $sevTable.Columns.Item(2).Width = $word.InchesToPoints(1.5)

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # =========================================================================
    # 2. RISK ASSESSMENT
    # =========================================================================
    $doc.Content.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    Add-Section '2. Risk Assessment' -Body (
        'The overall risk grade for this environment is {{ risk_grade }} with a ' +
        'risk score of {{ risk_score }}. The assessment identified ' +
        '{{ critical_issues }} critical issues and {{ high_risk_rules }} ' +
        'high-risk rules requiring immediate attention.'
    )

    $riskRange = $doc.Content.Paragraphs.Add().Range
    $riskTable = $doc.Tables.Add($riskRange, 4, 2)
    $riskTable.Borders.Enable = $true
    $riskItems = @(
        @('Best Practices Score',    '{{ best_practices_score }}'),
        @('Segmentation Score',      '{{ segmentation_score }}'),
        @('Lateral Movement Paths',  '{{ lateral_movement_paths }}'),
        @('Shadowed Rules',          '{{ shadowed_rule_count }}')
    )
    for ($i = 0; $i -lt 4; $i++) {
        Set-LabelValueRow -Row $riskTable.Rows.Item($i + 1) -Label $riskItems[$i][0] -Value $riskItems[$i][1]
    }

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # =========================================================================
    # 3. COMPLIANCE
    # =========================================================================
    Add-Section '3. Compliance Summary' -Body (
        'The following table summarises compliance status against applicable ' +
        'frameworks. This section is populated when compliance data is available ' +
        'in the Rampart export.'
    )

    $compRange = $doc.Content.Paragraphs.Add().Range
    $compTable = $doc.Tables.Add($compRange, 1, 6)
    $compTable.Borders.Enable = $true
    Set-HeaderRow -Row $compTable.Rows.Item(1) -Headers @('Framework','Score (%)','Status','Passed','Failed','Total')
    Add-MarkerRow -Table $compTable -MarkerName 'compliance' -Placeholders @('framework','percentage','status','passed','failed','total')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # =========================================================================
    # 4. DETAILED FINDINGS
    # =========================================================================
    $doc.Content.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    Add-Section '4. Detailed Findings' -Body (
        'This section presents all findings identified during the analysis, ' +
        'organised by severity.'
    )

    # 4.1 Critical
    Add-Section '4.1 Critical Findings' -Level 2 -Body '{{ critical_count }} critical finding(s) were identified.'

    $critRange = $doc.Content.Paragraphs.Add().Range
    $critTable = $doc.Tables.Add($critRange, 1, 7)
    $critTable.Borders.Enable = $true
    Set-HeaderRow -Row $critTable.Rows.Item(1) -Headers @('Rule','Device Group','Severity','Type','Description','Remediation','Risk')
    Add-MarkerRow -Table $critTable -MarkerName 'critical_findings' -Placeholders @('rule_name','device_group','severity','type','description','remediation','risk_score')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # 4.2 High
    Add-Section '4.2 High Findings' -Level 2 -Body '{{ high_count }} high-severity finding(s) were identified.'

    $highRange = $doc.Content.Paragraphs.Add().Range
    $highTable = $doc.Tables.Add($highRange, 1, 7)
    $highTable.Borders.Enable = $true
    Set-HeaderRow -Row $highTable.Rows.Item(1) -Headers @('Rule','Device Group','Severity','Type','Description','Remediation','Risk')
    Add-MarkerRow -Table $highTable -MarkerName 'high_findings' -Placeholders @('rule_name','device_group','severity','type','description','remediation','risk_score')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # 4.3 All
    Add-Section '4.3 All Findings' -Level 2 -Body '{{ total_findings }} finding(s) were identified in total.'

    $allRange = $doc.Content.Paragraphs.Add().Range
    $allTable = $doc.Tables.Add($allRange, 1, 7)
    $allTable.Borders.Enable = $true
    Set-HeaderRow -Row $allTable.Rows.Item(1) -Headers @('Rule','Device Group','Severity','Type','Description','Remediation','Risk')
    Add-MarkerRow -Table $allTable -MarkerName 'findings' -Placeholders @('rule_name','device_group','severity','type','description','remediation','risk_score')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # =========================================================================
    # 5. SHADOWED RULES
    # =========================================================================
    $doc.Content.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    Add-Section '5. Shadowed Rules' -Body (
        '{{ shadowed_rules_total }} shadowed rule(s) were detected. Shadowed rules ' +
        'are never matched because a broader rule higher in the policy takes precedence.'
    )

    $shadowRange = $doc.Content.Paragraphs.Add().Range
    $shadowTable = $doc.Tables.Add($shadowRange, 1, 6)
    $shadowTable.Borders.Enable = $true
    Set-HeaderRow -Row $shadowTable.Rows.Item(1) -Headers @('Rule','Shadowed By','Device Group','Severity','Description','Remediation')
    Add-MarkerRow -Table $shadowTable -MarkerName 'shadowed_rules' -Placeholders @('rule_name','shadowed_by','device_group','severity','description','remediation')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # =========================================================================
    # 6. DUPLICATE OBJECTS
    # =========================================================================
    Add-Section '6. Duplicate Objects' -Body (
        'The analysis identified {{ duplicate_address_count }} duplicate address ' +
        'object(s) and {{ duplicate_service_count }} duplicate service object(s).'
    )

    $dupRange = $doc.Content.Paragraphs.Add().Range
    $dupTable = $doc.Tables.Add($dupRange, 1, 5)
    $dupTable.Borders.Enable = $true
    Set-HeaderRow -Row $dupTable.Rows.Item(1) -Headers @('Type','Value','Count','Objects','Remediation')
    Add-MarkerRow -Table $dupTable -MarkerName 'duplicate_addresses' -Placeholders @('type','value','count','objects','remediation')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # =========================================================================
    # 7. NETWORK SEGMENTATION
    # =========================================================================
    $doc.Content.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    Add-Section '7. Network Segmentation' -Body (
        'Segmentation grade: {{ seg_grade }} (score: {{ seg_score }}). ' +
        'The analysis evaluated {{ seg_zone_count }} zone(s) with ' +
        '{{ seg_allowed_pairs }} allowed pair(s) and {{ seg_blocked_pairs }} blocked pair(s).'
    )

    # 7.1 Weak Segments
    Add-Section '7.1 Weak Segments' -Level 2

    $weakRange = $doc.Content.Paragraphs.Add().Range
    $weakTable = $doc.Tables.Add($weakRange, 1, 4)
    $weakTable.Borders.Enable = $true
    Set-HeaderRow -Row $weakTable.Rows.Item(1) -Headers @('Source Zone','Destination Zone','Openness','Remediation')
    Add-MarkerRow -Table $weakTable -MarkerName 'weak_segments' -Placeholders @('source_zone','dest_zone','openness','remediation')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # 7.2 Lateral Movement
    Add-Section '7.2 Lateral Movement' -Level 2 -Body '{{ lateral_movement_count }} rule(s) contribute to lateral movement risk.'

    $latRange = $doc.Content.Paragraphs.Add().Range
    $latTable = $doc.Tables.Add($latRange, 1, 5)
    $latTable.Borders.Enable = $true
    Set-HeaderRow -Row $latTable.Rows.Item(1) -Headers @('Rule','Severity','Source Zones','Destination Zones','Risk Factors')
    Add-MarkerRow -Table $latTable -MarkerName 'lateral_movement' -Placeholders @('rule_name','severity','source_zones','dest_zones','risk_factors')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # =========================================================================
    # 8. ADDITIONAL ANALYSIS
    # =========================================================================
    $doc.Content.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    Add-Section '8. Additional Analysis'

    # 8.1 Cleartext
    Add-Section '8.1 Cleartext Protocols' -Level 2 -Body '{{ cleartext_rule_count }} rule(s) permit cleartext protocols.'

    $clearRange = $doc.Content.Paragraphs.Add().Range
    $clearTable = $doc.Tables.Add($clearRange, 1, 4)
    $clearTable.Borders.Enable = $true
    Set-HeaderRow -Row $clearTable.Rows.Item(1) -Headers @('Rule','Protocol','Severity','Secure Alternative')
    Add-MarkerRow -Table $clearTable -MarkerName 'cleartext_rules' -Placeholders @('rule_name','protocol','severity','secure_alternative')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # 8.2 Stale Rules
    Add-Section '8.2 Stale Rules' -Level 2 -Body '{{ stale_rule_count }} stale rule(s) were identified.'

    $staleRange = $doc.Content.Paragraphs.Add().Range
    $staleTable = $doc.Tables.Add($staleRange, 1, 4)
    $staleTable.Borders.Enable = $true
    Set-HeaderRow -Row $staleTable.Rows.Item(1) -Headers @('Rule','Severity','Indicators','Disabled')
    Add-MarkerRow -Table $staleTable -MarkerName 'stale_rules' -Placeholders @('rule_name','severity','indicators','disabled')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # 8.3 Egress Risks
    Add-Section '8.3 Egress Risks' -Level 2 -Body '{{ egress_risk_count }} rule(s) present egress risk.'

    $egressRange = $doc.Content.Paragraphs.Add().Range
    $egressTable = $doc.Tables.Add($egressRange, 1, 4)
    $egressTable.Borders.Enable = $true
    Set-HeaderRow -Row $egressTable.Rows.Item(1) -Headers @('Rule','Severity','Risk Factors','Remediation')
    Add-MarkerRow -Table $egressTable -MarkerName 'egress_findings' -Placeholders @('rule_name','severity','risk_factors','remediation')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # 8.4 Decryption Gaps
    Add-Section '8.4 Decryption Gaps' -Level 2 -Body '{{ decryption_gap_count }} decryption gap(s) were identified.'

    $decRange = $doc.Content.Paragraphs.Add().Range
    $decTable = $doc.Tables.Add($decRange, 1, 4)
    $decTable.Borders.Enable = $true
    Set-HeaderRow -Row $decTable.Rows.Item(1) -Headers @('Rule','Severity','Reason','Remediation')
    Add-MarkerRow -Table $decTable -MarkerName 'decryption_gaps' -Placeholders @('rule_name','severity','reason','remediation')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # 8.5 Geo-IP
    Add-Section '8.5 Geo-IP Findings' -Level 2 -Body '{{ geo_ip_unrestricted_count }} rule(s) lack geo-IP restrictions.'

    $geoRange = $doc.Content.Paragraphs.Add().Range
    $geoTable = $doc.Tables.Add($geoRange, 1, 4)
    $geoTable.Borders.Enable = $true
    Set-HeaderRow -Row $geoTable.Rows.Item(1) -Headers @('Rule','Severity','Type','Remediation')
    Add-MarkerRow -Table $geoTable -MarkerName 'geo_ip_findings' -Placeholders @('rule_name','severity','type','remediation')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # 8.6 Rule Expiry
    Add-Section '8.6 Rule Expiry' -Level 2 -Body '{{ rule_expiry_count }} rule(s) have expiry concerns.'

    $expRange = $doc.Content.Paragraphs.Add().Range
    $expTable = $doc.Tables.Add($expRange, 1, 3)
    $expTable.Borders.Enable = $true
    Set-HeaderRow -Row $expTable.Rows.Item(1) -Headers @('Rule','Type','Detail')
    Add-MarkerRow -Table $expTable -MarkerName 'rule_expiry' -Placeholders @('rule_name','type','detail')

    $doc.Content.Paragraphs.Add().Range.InsertParagraphAfter()

    # =========================================================================
    # 9. RECOMMENDATIONS
    # =========================================================================
    $doc.Content.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    Add-Section '9. Recommendations' -Body 'Based on the findings of this audit, the following actions are recommended:'

    $recommendations = @(
        'Address all {{ critical_count }} critical finding(s) immediately.',
        'Review and remediate {{ high_count }} high-severity finding(s).',
        'Remove or consolidate {{ shadowed_rule_count }} shadowed rule(s) to simplify the policy.',
        'Eliminate {{ duplicate_address_count }} duplicate address object(s) and {{ duplicate_service_count }} duplicate service object(s).',
        'Review {{ stale_rule_count }} stale rule(s) for decommissioning.',
        'Replace cleartext protocols in {{ cleartext_rule_count }} rule(s) with encrypted alternatives.',
        'Improve network segmentation to reduce lateral movement paths (currently {{ lateral_movement_paths }}).'
    )
    foreach ($rec in $recommendations) {
        $p = $doc.Content.Paragraphs.Add()
        $p.Style = 'List Bullet'
        $p.Range.Text = $rec
        $p.Range.Font.Size = 11
        $p.Range.Font.Name = 'Calibri'
        $p.Range.InsertParagraphAfter()
    }

    # =========================================================================
    # APPENDIX
    # =========================================================================
    $doc.Content.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    Add-Section 'Appendix A: Methodology' -Body (
        'This audit was performed using Rampart, an automated firewall policy ' +
        'analysis platform. The configuration of type ''{{ config_type }}'' was ' +
        'analysed on {{ analysis_timestamp }}.'
    )

    Add-Section 'Appendix B: Best Practices' -Body (
        'Best practices assessment score: {{ best_practices_overall_score }} ' +
        '(grade: {{ best_practices_grade }}).'
    )

    # =========================================================================
    # Save
    # =========================================================================
    Write-Host "Saving template: $OutputPath"
    $doc.SaveAs2([ref]$OutputPath, [ref]$wdFormatXMLDocument)
    $doc.Close($false)

    Write-Host "Template created: $OutputPath"
}
catch {
    Write-Error "Error creating template: $_"
    if ($doc) { try { $doc.Close($false) } catch {} }
    exit 1
}
finally {
    if ($word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
