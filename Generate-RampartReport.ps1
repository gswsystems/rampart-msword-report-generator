<#
.SYNOPSIS
    Generates a Word (.docx) report from a Rampart JSON export using a Word template.

.DESCRIPTION
    Takes Rampart firewall audit JSON output and populates a Word template with
    findings, metrics, and analysis data. Uses Word COM automation.

    Templates use the same placeholder format as the Python rampart-report-generator:
      - {{ variable_name }}  for single-value placeholders
      - {{#table_name}} / {{/table_name}}  for repeating table rows

.PARAMETER JsonPath
    Path to the Rampart JSON export file.

.PARAMETER TemplatePath
    Path to the Word template (.docx) file.

.PARAMETER OutputPath
    Path for the generated report (.docx).

.PARAMETER ClientName
    Client/customer name for the report cover page.

.PARAMETER ClientContact
    Client contact person.

.PARAMETER AuditorName
    Auditor name.

.PARAMETER AuditorCompany
    Auditing company name.

.PARAMETER ReportTitle
    Report title (default: "Firewall Security Audit Report").

.PARAMETER ReportDate
    Report date string (default: today's date).

.PARAMETER Confidentiality
    Confidentiality marking (default: "CONFIDENTIAL").

.PARAMETER ListVariables
    List all available template variables from the JSON and exit.

.PARAMETER Visible
    Keep Word visible during generation (useful for debugging).

.EXAMPLE
    .\Generate-RampartReport.ps1 -JsonPath audit.json -TemplatePath template.docx -OutputPath report.docx
    .\Generate-RampartReport.ps1 -JsonPath audit.json -TemplatePath template.docx -OutputPath report.docx -ClientName "Acme Corp"
    .\Generate-RampartReport.ps1 -JsonPath audit.json -ListVariables

.NOTES
    Requires Microsoft Word installed (uses COM automation).
    Compatible with the template format from rampart-report-generator (Python).
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$JsonPath,

    [Parameter(Position = 1)]
    [string]$TemplatePath,

    [Parameter(Position = 2)]
    [string]$OutputPath,

    [string]$ClientName = '',
    [string]$ClientContact = '',
    [string]$AuditorName = '',
    [string]$AuditorCompany = '',
    [string]$ReportTitle = 'Firewall Security Audit Report',
    [string]$ReportDate = (Get-Date -Format 'yyyy-MM-dd'),
    [string]$Confidentiality = 'CONFIDENTIAL',

    [switch]$ListVariables,
    [switch]$Visible,
    [switch]$TrialWatermark
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# JSON loading
# ---------------------------------------------------------------------------

function Load-RampartJson {
    param([string]$Path)

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    $data = $raw | ConvertFrom-Json

    # Handle full report format (has 'analysis' wrapper) vs direct export
    if ($data.PSObject.Properties['analysis'] -and $data.PSObject.Properties['configuration']) {
        return $data.analysis
    }
    return $data
}

# ---------------------------------------------------------------------------
# Build flat variable dictionary
# ---------------------------------------------------------------------------

function Build-Variables {
    param(
        [psobject]$Analysis,
        [hashtable]$Overrides = @{}
    )

    $summary  = if ($Analysis.PSObject.Properties['summary'])  { $Analysis.summary }  else { [pscustomobject]@{} }
    $risk     = if ($Analysis.PSObject.Properties['risk_rating']) { $Analysis.risk_rating } else { [pscustomobject]@{} }
    $severity = if ($summary.PSObject.Properties['severity_breakdown']) { $summary.severity_breakdown } else { [pscustomobject]@{} }

    $v = @{}

    # --- Report metadata ---
    $v['report_date']       = if ($Overrides.ContainsKey('date') -and $Overrides['date']) { $Overrides['date'] } else { Get-Date -Format 'yyyy-MM-dd' }
    $v['report_title']      = if ($Overrides['title'])           { $Overrides['title'] }           else { 'Firewall Security Audit Report' }
    $v['client_name']       = if ($Overrides['client_name'])     { $Overrides['client_name'] }     else { '' }
    $v['client_contact']    = if ($Overrides['client_contact'])  { $Overrides['client_contact'] }  else { '' }
    $v['auditor_name']      = if ($Overrides['auditor_name'])    { $Overrides['auditor_name'] }    else { '' }
    $v['auditor_company']   = if ($Overrides['auditor_company']) { $Overrides['auditor_company'] } else { '' }
    $v['confidentiality']   = if ($Overrides['confidentiality']) { $Overrides['confidentiality'] } else { 'CONFIDENTIAL' }

    # --- Summary ---
    $v['total_rules']        = Get-JsonProp $summary 'total_rules_analyzed' 0
    $v['rules_with_issues']  = Get-JsonProp $summary 'rules_with_issues' 0
    $v['compliance_rate']    = [math]::Round((Get-JsonProp $summary 'compliance_rate' 0), 1)
    $v['config_type']        = Get-JsonProp $summary 'config_type' ''
    $v['analysis_timestamp'] = Get-JsonProp $summary 'timestamp' ''
    $v['device_group_count'] = Get-JsonProp $summary 'device_group_count' 0
    $dg = Get-JsonProp $summary 'device_groups' @()
    $v['device_groups']      = if ($dg -is [array]) { $dg -join ', ' } else { [string]$dg }

    # --- Severity counts ---
    $v['critical_count'] = Get-JsonProp $severity 'Critical' 0
    $v['high_count']     = Get-JsonProp $severity 'High' 0
    $v['medium_count']   = Get-JsonProp $severity 'Medium' 0
    $v['low_count']      = Get-JsonProp $severity 'Low' 0
    $v['total_findings'] = [int]$v['critical_count'] + [int]$v['high_count'] + [int]$v['medium_count'] + [int]$v['low_count']

    # --- Risk rating ---
    $v['risk_score']              = Get-JsonProp $risk 'score' 0
    $v['risk_grade']              = Get-JsonProp $risk 'grade' 'N/A'
    $v['best_practices_score']    = Get-JsonProp $risk 'best_practices_score' 0
    $v['segmentation_score']      = Get-JsonProp $risk 'segmentation_score' 0
    $v['critical_issues']         = Get-JsonProp $risk 'critical_issues' 0
    $v['high_risk_rules']         = Get-JsonProp $risk 'high_risk_rules' 0
    $v['shadowed_rule_count']     = Get-JsonProp $risk 'shadowed_rules' 0
    $v['lateral_movement_paths']  = Get-JsonProp $risk 'lateral_movement_paths' 0

    # --- Duplicate objects ---
    $dupes = if ($Analysis.PSObject.Properties['duplicate_objects']) { $Analysis.duplicate_objects } else { [pscustomobject]@{} }
    $v['duplicate_address_count'] = Get-JsonProp $dupes 'total_duplicate_addresses' 0
    $v['duplicate_service_count'] = Get-JsonProp $dupes 'total_duplicate_services' 0

    # --- Shadowed rules ---
    $shadowed = if ($Analysis.PSObject.Properties['shadowed_rules']) { $Analysis.shadowed_rules } else { @() }
    $v['shadowed_rules_total'] = if ($shadowed -is [array]) { $shadowed.Count } else { 0 }

    # --- Compliance ---
    $v['compliance_data_available'] = [bool]($Analysis.PSObject.Properties['compliance'])

    # --- Attack surface ---
    $v['attack_surface_available'] = [bool]($Analysis.PSObject.Properties['attack_surface'])

    # --- Best practices ---
    $bp = if ($Analysis.PSObject.Properties['best_practices']) { $Analysis.best_practices } else { $null }
    $v['best_practices_available'] = [bool]$bp
    if ($bp) {
        $v['best_practices_overall_score'] = Get-JsonProp $bp 'overall_score' 0
        $v['best_practices_grade']         = Get-JsonProp $bp 'grade' 'N/A'
    } else {
        $v['best_practices_overall_score'] = 0
        $v['best_practices_grade']         = 'N/A'
    }

    # --- Analyzer result counts ---
    $ruleExpiry = if ($Analysis.PSObject.Properties['rule_expiry']) { $Analysis.rule_expiry } else { [pscustomobject]@{} }
    $v['rule_expiry_count'] = (Get-JsonProp $ruleExpiry 'expired_schedule_count' 0) + (Get-JsonProp $ruleExpiry 'likely_temporary_count' 0)

    $v['cleartext_rule_count'] = Get-JsonProp (Get-JsonPropObj $Analysis 'cleartext_exposure') 'cleartext_rule_count' 0
    $v['geo_ip_unrestricted_count'] = Get-JsonProp (Get-JsonPropObj $Analysis 'geo_ip_exposure') 'unrestricted_count' 0
    $v['lateral_movement_count'] = Get-JsonProp (Get-JsonPropObj $Analysis 'lateral_movement') 'total_issues' 0

    $stale = if ($Analysis.PSObject.Properties['stale_rules']) { $Analysis.stale_rules } else { [pscustomobject]@{} }
    $v['stale_rule_count'] = (Get-JsonProp $stale 'stale_named_count' 0) + (Get-JsonProp $stale 'unused_object_rule_count' 0)

    $egress = if ($Analysis.PSObject.Properties['egress_filtering']) { $Analysis.egress_filtering } else { [pscustomobject]@{} }
    $eFindingsArr = if ($egress.PSObject.Properties['findings']) { $egress.findings } else { @() }
    $v['egress_risk_count'] = if ($eFindingsArr -is [array]) { $eFindingsArr.Count } else { 0 }

    $v['decryption_gap_count'] = Get-JsonProp (Get-JsonPropObj $Analysis 'decryption_policy') 'gaps_count' 0

    # --- Segmentation ---
    $segRoot = if ($Analysis.PSObject.Properties['segmentation_score']) { $Analysis.segmentation_score } else { [pscustomobject]@{} }
    $seg = if ($segRoot.PSObject.Properties['score']) { $segRoot.score } else { [pscustomobject]@{} }
    $v['seg_score']         = Get-JsonProp $seg 'segmentation_score' 0
    $v['seg_grade']         = Get-JsonProp $seg 'grade' 'N/A'
    $v['seg_zone_count']    = Get-JsonProp $seg 'zone_count' 0
    $v['seg_allowed_pairs'] = Get-JsonProp $seg 'allowed_pairs' 0
    $v['seg_blocked_pairs'] = Get-JsonProp $seg 'blocked_pairs' 0

    return $v
}

# ---------------------------------------------------------------------------
# Build table datasets
# ---------------------------------------------------------------------------

function Build-TableData {
    param([psobject]$Analysis)

    $tables = @{}

    # --- Findings ---
    $findingsRows = [System.Collections.ArrayList]::new()
    $rfList = if ($Analysis.PSObject.Properties['findings']) { $Analysis.findings } else { @() }
    foreach ($rf in $rfList) {
        $fList = if ($rf.PSObject.Properties['findings']) { $rf.findings } else { @() }
        foreach ($f in $fList) {
            [void]$findingsRows.Add(@{
                rule_name    = Get-JsonProp $rf 'rule_name' ''
                device_group = Get-JsonProp $rf 'device_group' ''
                severity     = Get-JsonProp $f 'severity' ''
                type         = Get-JsonProp $f 'type' ''
                description  = Get-JsonProp $f 'description' ''
                remediation  = Get-JsonProp $f 'remediation' ''
                risk_score   = Get-JsonProp $rf 'risk_score' 0
            })
        }
    }
    $tables['findings']          = $findingsRows
    $tables['critical_findings'] = @($findingsRows | Where-Object { $_['severity'] -eq 'Critical' })
    $tables['high_findings']     = @($findingsRows | Where-Object { $_['severity'] -eq 'High' })

    # --- Shadowed rules ---
    $shadowedList = if ($Analysis.PSObject.Properties['shadowed_rules']) { $Analysis.shadowed_rules } else { @() }
    $tables['shadowed_rules'] = @(foreach ($s in $shadowedList) {
        @{
            rule_name   = Get-JsonProp $s 'shadowed_rule_name' ''
            shadowed_by = Get-JsonProp $s 'shadowed_by_rule_name' ''
            device_group = Get-JsonProp $s 'device_group' ''
            severity    = Get-JsonProp $s 'severity' ''
            description = Get-JsonProp $s 'description' ''
            remediation = Get-JsonProp $s 'remediation' ''
        }
    })

    # --- Duplicate addresses ---
    $dupes = if ($Analysis.PSObject.Properties['duplicate_objects']) { $Analysis.duplicate_objects } else { [pscustomobject]@{} }
    $dupAddrs = if ($dupes.PSObject.Properties['duplicate_addresses']) { $dupes.duplicate_addresses } else { @() }
    $tables['duplicate_addresses'] = @(foreach ($d in $dupAddrs) {
        $names = if ($d.PSObject.Properties['object_names']) { $d.object_names } else { @() }
        @{
            type        = Get-JsonProp $d 'type' ''
            value       = Get-JsonProp $d 'value' ''
            count       = Get-JsonProp $d 'duplicate_count' 0
            objects     = ($names -join ', ')
            remediation = Get-JsonProp $d 'remediation' ''
        }
    })

    # --- Compliance frameworks ---
    $compliance = if ($Analysis.PSObject.Properties['compliance']) { $Analysis.compliance } else { $null }
    $compRows = [System.Collections.ArrayList]::new()
    if ($compliance) {
        foreach ($prop in $compliance.PSObject.Properties) {
            $val = $prop.Value
            if ($val -and $val.PSObject.Properties['compliance_percentage']) {
                [void]$compRows.Add(@{
                    framework  = $prop.Name
                    percentage = Get-JsonProp $val 'compliance_percentage' 0
                    status     = Get-JsonProp $val 'status' ''
                    passed     = Get-JsonProp $val 'passed_controls' 0
                    failed     = Get-JsonProp $val 'failed_controls' 0
                    total      = Get-JsonProp $val 'total_controls' 0
                })
            }
        }
    }
    $tables['compliance'] = $compRows

    # --- Lateral movement ---
    $latRoot = if ($Analysis.PSObject.Properties['lateral_movement']) { $Analysis.lateral_movement } else { [pscustomobject]@{} }
    $latRules = if ($latRoot.PSObject.Properties['lateral_movement_rules']) { $latRoot.lateral_movement_rules } else { @() }
    $tables['lateral_movement'] = @(foreach ($f in $latRules) {
        $sz = if ($f.PSObject.Properties['source_zones']) { $f.source_zones } else { @() }
        $dz = if ($f.PSObject.Properties['destination_zones']) { $f.destination_zones } else { @() }
        $rf2 = if ($f.PSObject.Properties['risk_factors']) { $f.risk_factors } else { @() }
        @{
            rule_name   = Get-JsonProp $f 'rule_name' ''
            severity    = Get-JsonProp $f 'severity' ''
            source_zones = ($sz -join ', ')
            dest_zones   = ($dz -join ', ')
            risk_factors = ($rf2 -join '; ')
        }
    })

    # --- Weak segments ---
    $segRoot = if ($Analysis.PSObject.Properties['segmentation_score']) { $Analysis.segmentation_score } else { [pscustomobject]@{} }
    $weakSegs = if ($segRoot.PSObject.Properties['weak_segments']) { $segRoot.weak_segments } else { @() }
    $tables['weak_segments'] = @(foreach ($s in $weakSegs) {
        @{
            source_zone = Get-JsonProp $s 'source_zone' ''
            dest_zone   = Get-JsonProp $s 'destination_zone' ''
            openness    = Get-JsonProp $s 'openness' ''
            remediation = Get-JsonProp $s 'remediation' ''
        }
    })

    # --- Egress findings ---
    $egressRoot = if ($Analysis.PSObject.Properties['egress_filtering']) { $Analysis.egress_filtering } else { [pscustomobject]@{} }
    $egressFindings = if ($egressRoot.PSObject.Properties['findings']) { $egressRoot.findings } else { @() }
    $tables['egress_findings'] = @(foreach ($f in $egressFindings) {
        $rf3 = if ($f.PSObject.Properties['risk_factors']) { $f.risk_factors } else { @() }
        @{
            rule_name    = Get-JsonProp $f 'rule_name' ''
            severity     = Get-JsonProp $f 'severity' ''
            risk_factors = ($rf3 -join '; ')
            remediation  = Get-JsonProp $f 'remediation' ''
        }
    })

    # --- Cleartext rules ---
    $clearRoot = if ($Analysis.PSObject.Properties['cleartext_exposure']) { $Analysis.cleartext_exposure } else { [pscustomobject]@{} }
    $clearRules = if ($clearRoot.PSObject.Properties['cleartext_rules']) { $clearRoot.cleartext_rules } else { @() }
    $tables['cleartext_rules'] = @(foreach ($f in $clearRules) {
        @{
            rule_name         = Get-JsonProp $f 'rule_name' ''
            protocol          = Get-JsonProp $f 'protocol' ''
            severity          = Get-JsonProp $f 'severity' ''
            secure_alternative = Get-JsonProp $f 'secure_alternative' ''
        }
    })

    # --- Stale rules ---
    $staleRoot = if ($Analysis.PSObject.Properties['stale_rules']) { $Analysis.stale_rules } else { [pscustomobject]@{} }
    $staleNamed = if ($staleRoot.PSObject.Properties['stale_named_rules']) { $staleRoot.stale_named_rules } else { @() }
    $staleUnused = if ($staleRoot.PSObject.Properties['unused_object_rules']) { $staleRoot.unused_object_rules } else { @() }
    $staleRows = [System.Collections.ArrayList]::new()
    foreach ($r in $staleNamed) {
        $ind = if ($r.PSObject.Properties['indicators']) { $r.indicators } else { @() }
        [void]$staleRows.Add(@{
            rule_name  = Get-JsonProp $r 'rule_name' ''
            severity   = Get-JsonProp $r 'severity' ''
            indicators = ($ind -join '; ')
            disabled   = if (Get-JsonProp $r 'disabled' $false) { 'Yes' } else { 'No' }
        })
    }
    foreach ($r in $staleUnused) {
        [void]$staleRows.Add(@{
            rule_name  = Get-JsonProp $r 'rule_name' ''
            severity   = 'High'
            indicators = 'References missing objects'
            disabled   = 'No'
        })
    }
    $tables['stale_rules'] = $staleRows

    # --- Decryption gaps ---
    $decRoot = if ($Analysis.PSObject.Properties['decryption_policy']) { $Analysis.decryption_policy } else { [pscustomobject]@{} }
    $decGaps = if ($decRoot.PSObject.Properties['gaps']) { $decRoot.gaps } else { @() }
    $tables['decryption_gaps'] = @(foreach ($g in $decGaps) {
        @{
            rule_name   = Get-JsonProp $g 'rule_name' ''
            severity    = Get-JsonProp $g 'severity' ''
            reason      = Get-JsonProp $g 'reason' ''
            remediation = Get-JsonProp $g 'remediation' ''
        }
    })

    # --- Geo-IP findings ---
    $geoRoot = if ($Analysis.PSObject.Properties['geo_ip_exposure']) { $Analysis.geo_ip_exposure } else { [pscustomobject]@{} }
    $geoUnrestricted = if ($geoRoot.PSObject.Properties['unrestricted_external_rules']) { $geoRoot.unrestricted_external_rules } else { @() }
    $geoMissing = if ($geoRoot.PSObject.Properties['missing_geo_block_rules']) { $geoRoot.missing_geo_block_rules } else { @() }
    $geoRows = [System.Collections.ArrayList]::new()
    foreach ($r in $geoUnrestricted) {
        [void]$geoRows.Add(@{
            rule_name   = Get-JsonProp $r 'rule_name' ''
            severity    = Get-JsonProp $r 'severity' ''
            type        = 'Unrestricted External'
            remediation = Get-JsonProp $r 'remediation' ''
        })
    }
    foreach ($r in $geoMissing) {
        [void]$geoRows.Add(@{
            rule_name   = Get-JsonProp $r 'rule_name' ''
            severity    = Get-JsonProp $r 'severity' ''
            type        = 'Missing Geo-Block'
            remediation = Get-JsonProp $r 'remediation' ''
        })
    }
    $tables['geo_ip_findings'] = $geoRows

    # --- Rule expiry ---
    $expiryRoot = if ($Analysis.PSObject.Properties['rule_expiry']) { $Analysis.rule_expiry } else { [pscustomobject]@{} }
    $expSchedule = if ($expiryRoot.PSObject.Properties['expired_schedule_rules']) { $expiryRoot.expired_schedule_rules } else { @() }
    $expTemp = if ($expiryRoot.PSObject.Properties['likely_temporary_rules']) { $expiryRoot.likely_temporary_rules } else { @() }
    $expiryRows = [System.Collections.ArrayList]::new()
    foreach ($r in $expSchedule) {
        $sched = Get-JsonProp $r 'schedule' ''
        $daysExp = Get-JsonProp $r 'days_expired' 0
        [void]$expiryRows.Add(@{
            rule_name = Get-JsonProp $r 'rule_name' ''
            type      = 'Expired Schedule'
            detail    = "Schedule: $sched, expired $daysExp days ago"
        })
    }
    foreach ($r in $expTemp) {
        [void]$expiryRows.Add(@{
            rule_name = Get-JsonProp $r 'rule_name' ''
            type      = 'Likely Temporary'
            detail    = Get-JsonProp $r 'reason' ''
        })
    }
    $tables['rule_expiry'] = $expiryRows

    return $tables
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-JsonProp {
    param($Obj, [string]$Name, $Default = $null)
    if ($null -eq $Obj) { return $Default }
    if ($Obj.PSObject.Properties[$Name]) {
        $val = $Obj.$Name
        if ($null -eq $val) { return $Default }
        return $val
    }
    return $Default
}

function Get-JsonPropObj {
    param($Obj, [string]$Name)
    if ($null -eq $Obj) { return [pscustomobject]@{} }
    if ($Obj.PSObject.Properties[$Name]) {
        $val = $Obj.$Name
        if ($null -eq $val) { return [pscustomobject]@{} }
        return $val
    }
    return [pscustomobject]@{}
}

# ---------------------------------------------------------------------------
# Word COM - Find & Replace in entire document
# ---------------------------------------------------------------------------

function Invoke-WordFindReplace {
    param(
        [object]$Document,
        [string]$Find,
        [string]$Replace
    )

    # wdReplaceAll = 2, wdFindContinue = 1
    $wdReplaceAll   = 2
    $wdFindContinue = 1

    # Replace in main body
    $range = $Document.Content
    $range.Find.ClearFormatting()
    $range.Find.Replacement.ClearFormatting()
    [void]$range.Find.Execute(
        $Find,                # FindText
        $false,               # MatchCase
        $false,               # MatchWholeWord
        $false,               # MatchWildcards
        $false,               # MatchSoundsLike
        $false,               # MatchAllWordForms
        $true,                # Forward
        $wdFindContinue,      # Wrap
        $false,               # Format
        $Replace,             # ReplaceWith
        $wdReplaceAll         # Replace
    )

    # Replace in headers and footers
    foreach ($section in $Document.Sections) {
        foreach ($header in $section.Headers) {
            $range = $header.Range
            $range.Find.ClearFormatting()
            $range.Find.Replacement.ClearFormatting()
            [void]$range.Find.Execute(
                $Find, $false, $false, $false, $false, $false, $true,
                $wdFindContinue, $false, $Replace, $wdReplaceAll
            )
        }
        foreach ($footer in $section.Footers) {
            $range = $footer.Range
            $range.Find.ClearFormatting()
            $range.Find.Replacement.ClearFormatting()
            [void]$range.Find.Execute(
                $Find, $false, $false, $false, $false, $false, $true,
                $wdFindContinue, $false, $Replace, $wdReplaceAll
            )
        }
    }
}

# ---------------------------------------------------------------------------
# Word COM - Process table markers
# ---------------------------------------------------------------------------

function Invoke-ProcessTables {
    param(
        [object]$Document,
        [hashtable]$TableData
    )

    foreach ($table in $Document.Tables) {
        $rowCount = $table.Rows.Count
        # Scan rows in reverse so deletions don't shift indices
        for ($rowIdx = $rowCount; $rowIdx -ge 1; $rowIdx--) {
            $row = $table.Rows.Item($rowIdx)
            $rowText = $row.Range.Text

            # Look for {{#table_name}} marker
            if ($rowText -match '\{\{#(\w+)\}\}') {
                $tableName = $Matches[1]
                $dataRows = $TableData[$tableName]

                if (-not $dataRows -or $dataRows.Count -eq 0) {
                    # No data - remove the marker row
                    $row.Delete()
                    continue
                }

                # This marker row is the template. We clone it for each data row.
                # First, collect the placeholder column names from the row cells.
                $colPlaceholders = @()
                for ($c = 1; $c -le $table.Columns.Count; $c++) {
                    $cellText = $table.Cell($rowIdx, $c).Range.Text
                    # Cell text ends with \r\a (end-of-cell + end-of-row markers)
                    $cellText = $cellText -replace '[\r\n\a]', ''
                    $cellText = $cellText.Trim()
                    # Remove table markers
                    $cellText = $cellText -replace '\{\{#\w+\}\}', ''
                    $cellText = $cellText -replace '\{\{/\w+\}\}', ''
                    $cellText = $cellText.Trim()
                    # Extract placeholder name: {{ name }}
                    if ($cellText -match '\{\{\s*(\w+)\s*\}\}') {
                        $colPlaceholders += $Matches[1]
                    } else {
                        $colPlaceholders += ''
                    }
                }

                # Insert data rows after the marker row, then delete the marker
                $insertAfterRow = $rowIdx
                foreach ($dataRow in $dataRows) {
                    # Add a new row after the current position
                    $newRow = $table.Rows.Add($table.Rows.Item([Math]::Min($insertAfterRow + 1, $table.Rows.Count)))
                    $insertAfterRow++

                    # Fill cells
                    for ($c = 0; $c -lt $colPlaceholders.Count; $c++) {
                        $ph = $colPlaceholders[$c]
                        if ($ph -and $dataRow.ContainsKey($ph)) {
                            $val = [string]$dataRow[$ph]
                        } else {
                            $val = ''
                        }
                        $cell = $newRow.Cells.Item($c + 1)
                        $cell.Range.Text = $val

                        # Copy font formatting from template row
                        $cell.Range.Font.Size = 9
                        $cell.Range.Font.Name = 'Calibri'
                    }
                }

                # Remove the marker row
                $row.Delete()
            }
        }
    }
}

# ---------------------------------------------------------------------------
# List variables mode
# ---------------------------------------------------------------------------

function Show-Variables {
    param(
        [hashtable]$Variables,
        [hashtable]$TableData
    )

    Write-Host ''
    Write-Host ('=' * 60)
    Write-Host 'TEMPLATE VARIABLES'
    Write-Host ('=' * 60)
    Write-Host ''
    Write-Host 'Use these in your template as {{ variable_name }}'
    Write-Host ''

    $groups = [ordered]@{
        'Report Metadata' = @('report_date','report_title','client_name','client_contact',
                              'auditor_name','auditor_company','confidentiality')
        'Summary'         = @('total_rules','rules_with_issues','compliance_rate','config_type',
                              'analysis_timestamp','device_group_count','device_groups')
        'Severity Counts' = @('critical_count','high_count','medium_count','low_count','total_findings')
        'Risk Rating'     = @('risk_score','risk_grade','best_practices_score','segmentation_score',
                              'critical_issues','high_risk_rules','shadowed_rule_count','lateral_movement_paths')
        'Duplicate Objects' = @('duplicate_address_count','duplicate_service_count')
        'Shadowed Rules'  = @('shadowed_rules_total')
        'Best Practices'  = @('best_practices_available','best_practices_overall_score','best_practices_grade')
        'Segmentation'    = @('seg_score','seg_grade','seg_zone_count','seg_allowed_pairs','seg_blocked_pairs')
        'Analyzer Counts' = @('rule_expiry_count','cleartext_rule_count','geo_ip_unrestricted_count',
                              'lateral_movement_count','stale_rule_count','egress_risk_count','decryption_gap_count')
    }

    foreach ($group in $groups.GetEnumerator()) {
        Write-Host "  $($group.Key):"
        foreach ($key in $group.Value) {
            $val = $Variables[$key]
            Write-Host "    {{ $key }} = $val"
        }
        Write-Host ''
    }

    Write-Host ('=' * 60)
    Write-Host 'TABLE DATA'
    Write-Host ('=' * 60)
    Write-Host ''
    Write-Host 'Use these in table rows with {{#table_name}} ... {{/table_name}}'
    Write-Host ''

    foreach ($entry in $TableData.GetEnumerator()) {
        $name = $entry.Key
        $rows = $entry.Value
        $count = if ($rows) { $rows.Count } else { 0 }
        if ($count -gt 0 -and $rows[0] -is [hashtable]) {
            $cols = ($rows[0].Keys | Sort-Object) -join ', '
            Write-Host "  {{#$name}} ($count rows)"
            Write-Host "    Columns: $cols"
        } else {
            Write-Host "  {{#$name}} ($count rows)"
        }
        Write-Host ''
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

# Resolve paths
$JsonPath = (Resolve-Path $JsonPath).Path

# Load JSON
$analysis = Load-RampartJson -Path $JsonPath

# Build data
$overrides = @{
    date            = $ReportDate
    title           = $ReportTitle
    client_name     = $ClientName
    client_contact  = $ClientContact
    auditor_name    = $AuditorName
    auditor_company = $AuditorCompany
    confidentiality = $Confidentiality
}

$variables = Build-Variables -Analysis $analysis -Overrides $overrides
$tableData = Build-TableData -Analysis $analysis

# List variables mode
if ($ListVariables) {
    Show-Variables -Variables $variables -TableData $tableData
    exit 0
}

# Validate remaining params
if (-not $TemplatePath -or -not $OutputPath) {
    Write-Error 'TemplatePath and OutputPath are required. Use -ListVariables to inspect the JSON without generating.'
    exit 1
}

$TemplatePath = (Resolve-Path $TemplatePath).Path
$OutputPath   = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

if (-not (Test-Path $TemplatePath)) {
    Write-Error "Template file not found: $TemplatePath"
    exit 1
}

# Launch Word
$word = $null
try {
    Write-Host "Opening Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$Visible
    $word.DisplayAlerts = 0  # wdAlertsNone

    # Open template
    Write-Host "Loading template: $TemplatePath"
    $doc = $word.Documents.Open($TemplatePath)

    # Step 1: Replace all {{ variable }} placeholders
    Write-Host "Replacing placeholders..."
    foreach ($entry in $variables.GetEnumerator()) {
        $placeholder = "{{ $($entry.Key) }}"
        $value = [string]$entry.Value

        # Word Find has a 255-char limit for replacement text
        if ($value.Length -gt 255) {
            $value = $value.Substring(0, 252) + '...'
        }

        Invoke-WordFindReplace -Document $doc -Find $placeholder -Replace $value
    }

    # Step 2: Process table markers
    Write-Host "Processing tables..."
    Invoke-ProcessTables -Document $doc -TableData $tableData

    # Step 3: Update TOC if present
    foreach ($toc in $doc.TablesOfContents) {
        $toc.Update()
    }

    # Step 4: Insert trial watermark if requested
    if ($TrialWatermark) {
        Write-Host "Adding trial watermark..."
        foreach ($section in $doc.Sections) {
            # Add watermark to the default (primary) header of each section
            $header = $section.Headers.Item(1)  # wdHeaderFooterPrimary = 1
            $headerRange = $header.Range

            # Insert a WordArt-style text watermark via a Shape
            $shape = $header.Shapes.AddTextEffect(
                0,                                       # msoTextEffect1 (plain)
                "Generated with Rampart Trial",          # Text
                "Calibri",                               # Font
                36,                                      # Size
                $false,                                  # Bold
                $false,                                  # Italic
                0,                                       # Left (centred later)
                0                                        # Top (centred later)
            )

            # Configure as a diagonal watermark
            $shape.TextEffect.NormalizedHeight = $false
            $shape.Line.Visible = $false
            $shape.Fill.Visible = $true
            $shape.Fill.ForeColor.RGB = 12632256  # Light grey (C0C0C0)
            $shape.Fill.Transparency = 0.75
            $shape.Rotation = 315                 # Diagonal (bottom-left to top-right)
            $shape.LockAspectRatio = $true

            # Centre on the page
            # wdWrapNone = 3, msoAnchorCenter values
            $shape.WrapFormat.Type = 3            # wdWrapNone — behind text
            $shape.RelativeHorizontalPosition = 0 # wdRelativeHorizontalPositionPage
            $shape.RelativeVerticalPosition = 0   # wdRelativeVerticalPositionPage
            $shape.Left = ($section.PageSetup.PageWidth - $shape.Width) / 2
            $shape.Top = ($section.PageSetup.PageHeight - $shape.Height) / 2
        }
    }

    # Step 5: Save as new document
    # wdFormatXMLDocument = 12 (.docx)
    Write-Host "Saving: $OutputPath"
    $doc.SaveAs2([ref]$OutputPath, [ref]12)
    $doc.Close($false)

    Write-Host "Report generated: $OutputPath"
}
catch {
    Write-Error "Error generating report: $_"
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
