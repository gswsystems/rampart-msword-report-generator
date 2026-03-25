# Rampart Word Report Generator — User Guide

Generate professional Microsoft Word reports from Rampart firewall audit exports using customisable `.docx` templates.

This tool reads the JSON output from a Rampart analysis, fills in a Word template with the audit data, and saves a finished `.docx` report. It uses Word COM automation, so Microsoft Word must be installed on the machine where reports are generated.

---

## Contents

1. [Quick Start](#1-quick-start)
2. [Installation & Requirements](#2-installation--requirements)
3. [Generating a Report](#3-generating-a-report)
4. [Inspecting Available Data](#4-inspecting-available-data)
5. [Creating Custom Templates](#5-creating-custom-templates)
   - [Placeholder Syntax](#51-placeholder-syntax)
   - [Table Markers](#52-table-markers)
   - [Step-by-Step: Building a Template from Scratch](#53-step-by-step-building-a-template-from-scratch)
   - [Step-by-Step: Modifying the Example Template](#54-step-by-step-modifying-the-example-template)
   - [Tips for Good Templates](#55-tips-for-good-templates)
6. [Calling from Rampart (QProcess)](#6-calling-from-rampart-qprocess)
7. [Variable Reference](#7-variable-reference)
   - [Report Metadata](#report-metadata)
   - [Summary](#summary)
   - [Severity Counts](#severity-counts)
   - [Risk Rating](#risk-rating)
   - [Duplicate Objects](#duplicate-objects)
   - [Shadowed Rules](#shadowed-rules)
   - [Best Practices](#best-practices)
   - [Segmentation](#segmentation)
   - [Analyser Counts](#analyser-counts)
8. [Table Reference](#8-table-reference)
   - [findings](#findings)
   - [critical_findings](#critical_findings)
   - [high_findings](#high_findings)
   - [shadowed_rules](#shadowed_rules)
   - [duplicate_addresses](#duplicate_addresses)
   - [compliance](#compliance)
   - [lateral_movement](#lateral_movement)
   - [weak_segments](#weak_segments)
   - [egress_findings](#egress_findings)
   - [cleartext_rules](#cleartext_rules)
   - [stale_rules](#stale_rules)
   - [decryption_gaps](#decryption_gaps)
   - [geo_ip_findings](#geo_ip_findings)
   - [rule_expiry](#rule_expiry)
9. [Troubleshooting](#9-troubleshooting)

---

## 1. Quick Start

```powershell
# 1. Generate the example template (one-time setup)
.\New-RampartTemplate.ps1

# 2. Generate a report from a Rampart JSON export
.\Generate-RampartReport.ps1 audit.json template.docx report.docx -ClientName "Acme Corp"
```

The output file `report.docx` is a standalone Word document ready for review and delivery.

---

## 2. Installation & Requirements

| Requirement        | Detail                                                   |
|--------------------|----------------------------------------------------------|
| Operating system   | Windows 10 / 11 or Windows Server 2016+                 |
| PowerShell         | 5.1 or later (ships with Windows)                       |
| Microsoft Word     | 2016, 2019, 2021, or Microsoft 365 (desktop)            |
| Rampart export     | A `.json` file exported from Rampart's analysis engine   |

No additional PowerShell modules or packages are needed. The script uses the Word COM interop that is available whenever Word is installed.

### Execution policy

If you have not run PowerShell scripts on this machine before, you may need to allow script execution:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Alternatively, invoke the script with the bypass flag:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Generate-RampartReport.ps1 ...
```

---

## 3. Generating a Report

### Basic usage

```powershell
.\Generate-RampartReport.ps1 <JsonPath> <TemplatePath> <OutputPath>
```

### All parameters

| Parameter | Req. | Default | Description |
|---|---|---|---|
| `-JsonPath` | Yes | — | Rampart JSON export path |
| `-TemplatePath` | Yes* | — | Word template path |
| `-OutputPath` | Yes* | — | Output report path |
| `-ClientName` | No | *(empty)* | Client or customer name |
| `-ClientContact` | No | *(empty)* | Client contact person |
| `-AuditorName` | No | *(empty)* | Auditor name |
| `-AuditorCompany` | No | *(empty)* | Auditing company name |
| `-ReportTitle` | No | `Firewall Sec...` | Cover page and footer title |
| `-ReportDate` | No | Today (`yyyy-MM-dd`) | Report date string |
| `-Confidentiality` | No | `CONFIDENTIAL` | Confidentiality marking |
| `-ListVariables` | No | — | Print variables and exit |
| `-Visible` | No | — | Keep Word visible (debug) |

*Not required when using `-ListVariables`.

### Examples

```powershell
# Minimal — just JSON, template, and output
.\Generate-RampartReport.ps1 audit.json template.docx report.docx

# With metadata
.\Generate-RampartReport.ps1 audit.json template.docx report.docx `
    -ClientName "Acme Corp" `
    -ClientContact "Jane Smith" `
    -AuditorName "John Doe" `
    -AuditorCompany "SecureAudit Ltd" `
    -ReportDate "2026-03-25" `
    -Confidentiality "STRICTLY CONFIDENTIAL"

# Custom title
.\Generate-RampartReport.ps1 audit.json template.docx report.docx `
    -ReportTitle "Q1 2026 Firewall Review" `
    -ClientName "Globex"

# Debug mode — watch Word fill in the template
.\Generate-RampartReport.ps1 audit.json template.docx report.docx -Visible
```

### What the script does

1. Reads the Rampart JSON export and extracts all analysis data.
2. Opens the template in Word (hidden by default).
3. Replaces every `{{ variable }}` placeholder with its value.
4. Expands table markers — each `{{#table_name}}` row is cloned once per data record.
5. Updates the Table of Contents if one is present.
6. Saves the result as a new `.docx` file and closes Word.

---

## 4. Inspecting Available Data

Before building a template, check what data is available in your JSON file:

```powershell
.\Generate-RampartReport.ps1 audit.json -ListVariables
```

This prints every variable and its current value, plus all table datasets with their column names and row counts. Use this output to decide which placeholders to include in your template.

Example output:

```
============================================================
TEMPLATE VARIABLES
============================================================

Use these in your template as {{ variable_name }}

  Report Metadata:
    {{ report_date }} = 2026-03-25
    {{ report_title }} = Firewall Security Audit Report
    {{ client_name }} =
    ...

  Severity Counts:
    {{ critical_count }} = 12
    {{ high_count }} = 47
    ...

============================================================
TABLE DATA
============================================================

Use these in table rows with {{#table_name}} ... {{/table_name}}

  {{#findings}} (193 rows)
    Columns: description, device_group, remediation, risk_score, rule_name, severity, type

  {{#critical_findings}} (12 rows)
    Columns: description, device_group, remediation, risk_score, rule_name, severity, type
  ...
```

---

## 5. Creating Custom Templates

A template is a standard `.docx` file that you create and edit in Microsoft Word. You insert special placeholder tokens into the document text; the script replaces them with real data at generation time.

### 5.1 Placeholder Syntax

Single-value placeholders use double curly braces:

```
{{ variable_name }}
```

Place these anywhere in your document — body text, headers, footers, table cells, text boxes. The script performs a global find-and-replace, so every occurrence is substituted.

**Examples in running text:**

> This report was prepared for **{{ client_name }}** on **{{ report_date }}**.
> A total of **{{ total_rules }}** firewall rules were analysed, of which
> **{{ rules_with_issues }}** had one or more findings.

**Examples in a table cell:**

| Metric         | Value                 |
|----------------|-----------------------|
| Total Rules    | {{ total_rules }}     |
| Risk Grade     | {{ risk_grade }}      |

The spaces inside the braces are required: `{{ name }}`, not `{{name}}`.

### 5.2 Table Markers

For repeating data (like a list of findings or compliance frameworks), use table markers inside a Word table. The marker row acts as a template — it is cloned once for each data record and then removed.

#### Structure

A Word table with markers has three parts:

1. **Header row** — static column headers (styled however you like).
2. **Marker row** — contains `{{#table_name}}` in the first cell and `{{ column }}` placeholders in subsequent cells. This row also includes `{{/table_name}}` in the first cell (or the last cell) to close the block.
3. The script removes the marker row and inserts one row per data record in its place.

#### Example

Create a table in Word like this:

| Rule | Severity | Description | Remediation |
|------|----------|-------------|-------------|
| {{#critical_findings}} {{ rule_name }} | {{ severity }} | {{ description }} | {{ remediation }} {{/critical_findings}} |

When the report is generated with 3 critical findings, the output becomes:

| Rule | Severity | Description | Remediation |
|------|----------|-------------|-------------|
| Allow-All-DMZ | Critical | Overly permissive rule allows... | Restrict source addresses to... |
| Any-Any-Trust | Critical | Rule permits all traffic... | Remove or scope this rule... |
| No-Profile-WAN | Critical | Missing security profiles... | Attach threat prevention... |

If there are no matching records, the marker row is simply removed, leaving only the header.

#### Marker placement

The start marker `{{#table_name}}` and end marker `{{/table_name}}` must appear in the **same row**. The typical pattern is:

- **First cell:** `{{#table_name}} {{ first_column }}`
- **Middle cells:** `{{ column_2 }}`, `{{ column_3 }}`, etc.
- **Last cell:** `{{ last_column }} {{/table_name}}`

Or you can put markers and placeholders in separate cells — the script strips out the marker tags and only uses the `{{ column }}` placeholders to fill data.

The example template (`New-RampartTemplate.ps1`) uses a three-row approach for clarity:

- **Row 1 (start marker):** `{{#table_name}}` in the first cell, remaining cells empty.
- **Row 2 (data row):** `{{ col1 }}`, `{{ col2 }}`, etc.
- **Row 3 (end marker):** `{{/table_name}}` in the first cell, remaining cells empty.

Both approaches work. Use whichever is clearer for your template.

### 5.3 Step-by-Step: Building a Template from Scratch

1. **Open a new document in Word.** Set up your page size, margins, fonts, and any branding (logo, colours, headers/footers).

2. **Add a cover page.** Type your title using `{{ report_title }}`, client name using `{{ client_name }}`, date using `{{ report_date }}`, and any other metadata placeholders.

3. **Add an executive summary section.** Write your narrative text and embed placeholders where numbers or values should appear:

   > The audit analysed {{ total_rules }} rules and found {{ total_findings }}
   > issues, including {{ critical_count }} critical and {{ high_count }} high-severity findings.
   > The overall risk grade is {{ risk_grade }}.

4. **Add data tables.** Insert a Word table, style the header row, then add a marker row beneath it. For example, a findings table:

   - Insert a table with columns: Rule, Severity, Type, Description, Remediation.
   - Style the first row as the header (bold, shaded, etc.).
   - In the second row, type:
     - Cell 1: `{{#findings}} {{ rule_name }}`
     - Cell 2: `{{ severity }}`
     - Cell 3: `{{ type }}`
     - Cell 4: `{{ description }}`
     - Cell 5: `{{ remediation }} {{/findings}}`

5. **Add any additional sections** using the variables and tables listed in the reference below.

6. **Optionally insert a Table of Contents** (References > Table of Contents in Word). The script will update it automatically after filling in the data.

7. **Save as `.docx`.** This is your template — do not save as `.doc` or `.dotx`.

### 5.4 Step-by-Step: Modifying the Example Template

1. **Generate the example template:**

   ```powershell
   .\New-RampartTemplate.ps1
   ```

2. **Open `template.docx` in Word.**

3. **Customise the layout:**
   - Add your company logo to the cover page or header.
   - Change fonts, colours, and table styles to match your branding.
   - Rearrange, remove, or duplicate sections as needed.

4. **Remove sections you don't need.** If you don't want the Geo-IP section, just delete that heading, paragraph, and table from the document. The script only replaces placeholders that exist in the template.

5. **Add new narrative.** You can write any text around the placeholders. The script only touches `{{ ... }}` tokens and `{{# ... }}` table markers — everything else is left exactly as-is.

6. **Save.** Your customised template is ready to use.

### 5.5 Tips for Good Templates

- **Formatting is preserved.** If you make `{{ risk_grade }}` bold and red in the template, the replaced value will also be bold and red.

- **Unknown placeholders are left as-is.** If you mistype a variable name (`{{ totl_rules }}`), it will appear literally in the output. Use `-ListVariables` to verify names.

- **Long values are truncated to 255 characters** in single-value replacements due to a Word Find limitation. Table cell values are not affected by this limit.

- **One template, many reports.** The same template works for any Rampart JSON export. Sections with zero data (e.g., no critical findings) will have their table marker rows removed, leaving just the header row and any surrounding text.

- **Headers and footers are processed.** You can put `{{ client_name }}`, `{{ report_date }}`, or `{{ confidentiality }}` in your document headers and footers.

- **Table of Contents is updated automatically.** If your template includes a Word TOC field, the script calls `Update` on it after all replacements are done. The user may still need to confirm the update when first opening the document in Word.

- **Test with `-Visible`.** When developing a template, use the `-Visible` flag to watch Word process the document in real time. This makes it easy to spot layout issues.

---

## 6. Calling from Rampart (QProcess)

The script is designed to be invoked from Rampart's Qt application via `QProcess`. Here is the typical integration pattern:

```cpp
QProcess *process = new QProcess(this);
QStringList args;

args << "-ExecutionPolicy" << "Bypass"
     << "-File" << reportGeneratorPath  // path to Generate-RampartReport.ps1
     << jsonExportPath                  // path to the exported JSON
     << templatePath                    // path to the .docx template
     << outputPath                      // desired output .docx path
     << "-ClientName" << clientName
     << "-AuditorName" << auditorName
     << "-AuditorCompany" << companyName;

process->start("powershell.exe", args);
```

The script:
- Writes progress messages to **stdout** (`Opening Word...`, `Replacing placeholders...`, `Report generated: <path>`).
- Writes errors to **stderr**.
- Exits with code **0** on success, **1** on failure.

Connect to `QProcess::finished` to check the exit code and `QProcess::readAllStandardOutput` / `readAllStandardError` for status messages.

---

## 7. Variable Reference

All single-value placeholders available for use in templates. Values shown are representative examples.

### Report Metadata

These are set via command-line parameters, not from the JSON data.

| Variable | Type | Example | Description |
|---|---|---|---|
| `report_date` | string | `2026-03-25` | Report date (default: today) |
| `report_title` | string | `Firewall Sec...` | Report title |
| `client_name` | string | `Acme Corp` | Client / customer name |
| `client_contact` | string | `Jane Smith` | Client contact person |
| `auditor_name` | string | `John Doe` | Auditor name |
| `auditor_company` | string | `SecureAudit Ltd` | Auditing company |
| `confidentiality` | string | `CONFIDENTIAL` | Confidentiality marking |

### Summary

Derived from the `summary` section of the Rampart JSON export.

| Variable | Type | Example | Description |
|---|---|---|---|
| `total_rules` | integer | `342` | Total firewall rules analysed |
| `rules_with_issues` | integer | `87` | Rules with at least one finding |
| `compliance_rate` | float | `74.6` | Overall compliance rate (%) |
| `config_type` | string | `panorama` | `standalone` or `panorama` |
| `analysis_timestamp` | string | `2026-03-25T...` | ISO 8601 analysis timestamp |
| `device_group_count` | integer | `3` | Number of device groups |
| `device_groups` | string | `DG-Corp, ...` | Comma-separated device groups |

### Severity Counts

Counts of findings by severity level, derived from `summary.severity_breakdown`.

| Variable | Type | Example | Description |
|---|---|---|---|
| `critical_count` | integer | `12` | Critical-severity findings |
| `high_count` | integer | `47` | High-severity findings |
| `medium_count` | integer | `63` | Medium-severity findings |
| `low_count` | integer | `71` | Low-severity findings |
| `total_findings` | integer | `193` | Sum of all severity counts |

### Risk Rating

Derived from the `risk_rating` section of the JSON export.

| Variable | Type | Example | Description |
|---|---|---|---|
| `risk_score` | integer | `62` | Risk score (0--100, lower is better) |
| `risk_grade` | string | `D` | Letter grade (A--F) |
| `best_practices_score` | integer | `58` | Best practices score (0--100) |
| `segmentation_score` | integer | `45` | Segmentation score (0--100) |
| `critical_issues` | integer | `12` | Critical issues contributing to risk |
| `high_risk_rules` | integer | `23` | Count of high-risk rules |
| `shadowed_rule_count` | integer | `8` | Shadowed (unreachable) rules |
| `lateral_movement_paths` | integer | `5` | Lateral movement paths found |

### Duplicate Objects

Derived from the `duplicate_objects` section.

| Variable | Type | Example | Description |
|---|---|---|---|
| `duplicate_address_count` | integer | `14` | Duplicate address objects |
| `duplicate_service_count` | integer | `3` | Duplicate service objects |

### Shadowed Rules

| Variable | Type | Example | Description |
|---|---|---|---|
| `shadowed_rules_total` | integer | `8` | Total shadowed rules |

### Best Practices

Derived from the `best_practices` section.

| Variable | Type | Example | Description |
|---|---|---|---|
| `best_practices_available` | boolean | `True` | Whether data is present |
| `best_practices_overall_score` | integer | `58` | Overall score (0--100) |
| `best_practices_grade` | string | `C` | Letter grade |

### Segmentation

Derived from the `segmentation_score.score` section.

| Variable | Type | Example | Description |
|---|---|---|---|
| `seg_score` | integer | `45` | Effectiveness score (0--100) |
| `seg_grade` | string | `D` | Segmentation letter grade |
| `seg_zone_count` | integer | `6` | Number of security zones |
| `seg_allowed_pairs` | integer | `12` | Allowed zone-to-zone pairs |
| `seg_blocked_pairs` | integer | `18` | Blocked zone-to-zone pairs |

### Analyser Counts

Summary counts from individual analysers. Useful for section introductions (e.g., "12 rule(s) permit cleartext protocols").

| Variable | Type | Example | Description |
|---|---|---|---|
| `rule_expiry_count` | integer | `4` | Expired + temporary rules |
| `cleartext_rule_count` | integer | `12` | Cleartext protocol rules |
| `geo_ip_unrestricted_count` | integer | `6` | Rules without geo-IP limits |
| `lateral_movement_count` | integer | `5` | Lateral movement risk rules |
| `stale_rule_count` | integer | `9` | Stale or unused rules |
| `egress_risk_count` | integer | `7` | Rules with egress risk |
| `decryption_gap_count` | integer | `3` | SSL/TLS decryption gaps |

---

## 8. Table Reference

All table datasets available for use with `{{#table_name}}` markers. Each table lists its column names — use these as `{{ column }}` placeholders inside the marker row.

---

### findings

All findings across all severity levels.

| Column | Type | Description |
|---|---|---|
| `rule_name` | string | Firewall rule name |
| `device_group` | string | Device group the rule belongs to |
| `severity` | string | `Critical`, `High`, `Medium`, or `Low` |
| `type` | string | Finding type (e.g., `overly_permissive`) |
| `description` | string | Explanation of the issue |
| `remediation` | string | Recommended fix |
| `risk_score` | integer | Numeric risk score for the rule |

### critical_findings

Subset of `findings` where `severity` is `Critical`. Same columns as `findings`.

### high_findings

Subset of `findings` where `severity` is `High`. Same columns as `findings`.

---

### shadowed_rules

Rules that are never matched because a broader rule higher in the policy takes precedence.

| Column         | Type   | Description                                  |
|----------------|--------|----------------------------------------------|
| `rule_name`    | string | Name of the shadowed rule                    |
| `shadowed_by`  | string | Name of the rule that shadows it             |
| `device_group` | string | Device group                                 |
| `severity`     | string | Severity level                               |
| `description`  | string | Explanation of the shadowing                 |
| `remediation`  | string | Recommended action                           |

---

### duplicate_addresses

Duplicate address objects identified in the configuration.

| Column        | Type    | Description                                        |
|---------------|---------|----------------------------------------------------|
| `type`        | string  | Object type (e.g., `ip-netmask`, `fqdn`)           |
| `value`       | string  | The duplicated value                               |
| `count`       | integer | Number of objects sharing this value               |
| `objects`     | string  | Comma-separated list of object names               |
| `remediation` | string  | Recommended consolidation action                   |

---

### compliance

Compliance assessment results per framework.

| Column       | Type    | Description                                    |
|--------------|---------|------------------------------------------------|
| `framework`  | string  | Framework name (e.g., `CIS`, `PCI-DSS`)        |
| `percentage` | float   | Compliance percentage                          |
| `status`     | string  | Overall status                                 |
| `passed`     | integer | Number of controls passed                      |
| `failed`     | integer | Number of controls failed                      |
| `total`      | integer | Total controls evaluated                       |

---

### lateral_movement

Rules that facilitate lateral movement between network zones.

| Column         | Type   | Description                                        |
|----------------|--------|----------------------------------------------------|
| `rule_name`    | string | Rule name                                          |
| `severity`     | string | Severity level                                     |
| `source_zones` | string | Comma-separated source zones                       |
| `dest_zones`   | string | Comma-separated destination zones                  |
| `risk_factors` | string | Semicolon-separated risk factors                   |

---

### weak_segments

Zone pairs with insufficient segmentation.

| Column        | Type   | Description                                    |
|---------------|--------|------------------------------------------------|
| `source_zone` | string | Source security zone                           |
| `dest_zone`   | string | Destination security zone                      |
| `openness`    | string | Degree of openness between zones               |
| `remediation` | string | Recommended action to improve segmentation     |

---

### egress_findings

Rules that present egress (outbound) risk.

| Column         | Type   | Description                                 |
|----------------|--------|---------------------------------------------|
| `rule_name`    | string | Rule name                                   |
| `severity`     | string | Severity level                              |
| `risk_factors` | string | Semicolon-separated risk factors            |
| `remediation`  | string | Recommended action                          |

---

### cleartext_rules

Rules permitting cleartext (unencrypted) protocols.

| Column               | Type   | Description                              |
|----------------------|--------|------------------------------------------|
| `rule_name`          | string | Rule name                                |
| `protocol`           | string | Cleartext protocol (e.g., `FTP`, `HTTP`) |
| `severity`           | string | Severity level                           |
| `secure_alternative` | string | Recommended encrypted alternative        |

---

### stale_rules

Rules that appear to be stale, unused, or obsolete.

| Column       | Type   | Description                                          |
|--------------|--------|------------------------------------------------------|
| `rule_name`  | string | Rule name                                            |
| `severity`   | string | Severity level                                       |
| `indicators` | string | Semicolon-separated staleness indicators             |
| `disabled`   | string | `Yes` if the rule is disabled, `No` otherwise        |

---

### decryption_gaps

Gaps in SSL/TLS decryption policy coverage.

| Column        | Type   | Description                              |
|---------------|--------|------------------------------------------|
| `rule_name`   | string | Rule name                                |
| `severity`    | string | Severity level                           |
| `reason`      | string | Why this is a gap                        |
| `remediation` | string | Recommended action                       |

---

### geo_ip_findings

Rules lacking geographic IP restrictions or missing geo-block policies.

| Column        | Type   | Description                                               |
|---------------|--------|-----------------------------------------------------------|
| `rule_name`   | string | Rule name                                                 |
| `severity`    | string | Severity level                                            |
| `type`        | string | `Unrestricted External` or `Missing Geo-Block`            |
| `remediation` | string | Recommended action                                        |

---

### rule_expiry

Rules with expired schedules or that appear to be temporary.

| Column      | Type   | Description                                                 |
|-------------|--------|-------------------------------------------------------------|
| `rule_name` | string | Rule name                                                   |
| `type`      | string | `Expired Schedule` or `Likely Temporary`                    |
| `detail`    | string | Schedule and expiry info, or reason for flagging            |

---

## 9. Troubleshooting

### "Word.Application" COM object could not be created

Word is not installed, or the COM registration is broken. Verify that Word opens normally, then try:

```powershell
# Re-register Word COM (run as Administrator)
& "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" /regserver
```

### Script hangs or Word becomes unresponsive

A Word dialog may be waiting for input (e.g., a macro security prompt or a document recovery dialog). Run with `-Visible` to see what Word is showing. Close any dialogs, then re-run.

If Word processes are left behind after an error:

```powershell
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
```

### Placeholders appear literally in the output

- Check that the placeholder name matches exactly (use `-ListVariables` to verify).
- Ensure there are spaces inside the braces: `{{ name }}` not `{{name}}`.
- If you typed the placeholder in Word and then reformatted parts of it (e.g., bolded just the variable name), Word may have split the text across multiple XML runs internally. The simplest fix is to delete the placeholder text and retype it in one go without changing formatting mid-placeholder.

### Table rows are not populated

- Verify the table has a marker row containing `{{#table_name}}`.
- Check that column placeholder names match the table reference above.
- Use `-ListVariables` to confirm the table has data (row count > 0).

### Output file is locked / cannot be saved

Another process (or a previous failed run) may have the file open. Close Word and any open instances of the output file, then retry.

### Exit codes

| Code | Meaning                                                   |
|------|-----------------------------------------------------------|
| 0    | Report generated successfully                             |
| 1    | Error (missing parameters, file not found, Word failure)  |
