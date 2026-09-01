Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web



# ==========================================
# AIRTABLE SETTINGS
# ==========================================

$airtableFormUrl = "https://airtable.com/appybOSI4sqIAk36T/pagFt8aEkGJbcE2oO/form"

# Saves which analysts are selected
$analystSettingsFile = "$PSScriptRoot\Stellar-Airtable-Analysts.json"

# ==========================================
# ANALYST LIST
# ==========================================

$analysts = @(
    "Bruce Jamail"
    "Ruth A Nolan"
    "Bashar Al Qaraghuli"
    "Travis Fletcher"
    "Shloka Jain"
    "Chi-Heng Chan"
)

# ==========================================
# CREATE WINDOW
# ==========================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Stellar → Airtable"
$form.Size = New-Object System.Drawing.Size(700, 650)
$form.StartPosition = "CenterScreen"


# ==========================================
# INSTRUCTIONS
# ==========================================

$label = New-Object System.Windows.Forms.Label
$label.Text = "Paste the Stellar alert information below:"
$label.Location = New-Object System.Drawing.Point(20, 20)
$label.Size = New-Object System.Drawing.Size(500, 25)
$form.Controls.Add($label) 

# ==========================================
# TEXT BOX
# ==========================================

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(20, 50)
$textBox.Size = New-Object System.Drawing.Size(640, 360)
$textBox.Multiline = $true
$textBox.ScrollBars = "Vertical"
$form.Controls.Add($textBox)

# ==========================================
# ANALYST CHECKLIST
# ==========================================

$analystLabel = New-Object System.Windows.Forms.Label
$analystLabel.Text = "Analysts on Shift:"
$analystLabel.Location = New-Object System.Drawing.Point(20, 420)
$analystLabel.Size = New-Object System.Drawing.Size(200, 25)
$analystLabel.ForeColor = [System.Drawing.Color]::FromArgb(90, 160, 220)
$form.Controls.Add($analystLabel)

$analystChecklist = New-Object System.Windows.Forms.CheckedListBox
$analystChecklist.Location = New-Object System.Drawing.Point(20, 445)

# Stretch across almost the full window
# Height is sized for about 3 analyst rows
$analystChecklist.Size = New-Object System.Drawing.Size(640, 65)

# After 3 rows, continue into another column
$analystChecklist.MultiColumn = $true
$analystChecklist.ColumnWidth = 200

$analystChecklist.CheckOnClick = $true
$analystChecklist.BackColor = [System.Drawing.Color]::FromArgb(38, 42, 46)
$analystChecklist.ForeColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
$analystChecklist.BorderStyle = "FixedSingle"
$form.Controls.Add($analystChecklist)

# Load saved analyst selections
$savedAnalysts = $null

if (Test-Path $analystSettingsFile) {
    try {
        $savedAnalysts = Get-Content $analystSettingsFile -Raw | ConvertFrom-Json
    }
    catch {
        $savedAnalysts = $null
    }
}

foreach ($analyst in $analysts) {

    # Default to unchecked
    $shouldBeChecked = $false

    # Use saved setting if one exists
    if ($null -ne $savedAnalysts) {
        if ($savedAnalysts.PSObject.Properties.Name -contains $analyst) {
            $shouldBeChecked = [bool]$savedAnalysts.($analyst)
        }
    }

    [void]$analystChecklist.Items.Add(
        $analyst,
        $shouldBeChecked
    )
}


# ==========================================
# FILL AIRTABLE BUTTON
# ==========================================

$button = New-Object System.Windows.Forms.Button
$button.Text = "Fill Airtable"
$button.Location = New-Object System.Drawing.Point(250, 535)
$button.Size = New-Object System.Drawing.Size(180, 45)
$form.Controls.Add($button)

# ==========================================
# STYLING
# ==========================================

# Window
$form.BackColor = [System.Drawing.Color]::FromArgb(32, 34, 37)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Instruction label
$label.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$label.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)

# Main text box
$textBox.BackColor = [System.Drawing.Color]::FromArgb(45, 48, 52)
$textBox.ForeColor = [System.Drawing.Color]::White
$textBox.BorderStyle = "FixedSingle"
$textBox.Font = New-Object System.Drawing.Font("Consolas", 10)

# Fill Airtable button
$button.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$button.ForeColor = [System.Drawing.Color]::White
$button.FlatStyle = "Flat"
$button.FlatAppearance.BorderSize = 0
$button.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$button.Cursor = [System.Windows.Forms.Cursors]::Hand 

$button.Add_MouseEnter({
    $button.BackColor = [System.Drawing.Color]::FromArgb(0, 140, 240)
})

$button.Add_MouseLeave({
    $button.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
})



# ==========================================
# BUTTON ACTION
# ==========================================

$button.Add_Click({

    $text = $textBox.Text

    # -----------------------------
    # JSON DATA
    # -----------------------------

    $jsonText = ""
    $jsonSourceIP = ""
    $jsonDestinationIP = ""
    $jsonHostIP = ""

    # Extract everything from the first { through the last }
    $firstBrace = $text.IndexOf("{")
    $lastBrace = $text.LastIndexOf("}")

    if ($firstBrace -ge 0 -and $lastBrace -gt $firstBrace) {
        $jsonText = $text.Substring(
            $firstBrace,
            $lastBrace - $firstBrace + 1
        )
    }

    # Source IP
    if ($jsonText -match '(?i)"srcip"\s*:\s*"([^"]+)"') {
        $jsonSourceIP = $matches[1].Trim()
    }

    # Destination IP
    if ($jsonText -match '(?i)"dstip"\s*:\s*"([^"]+)"') {
        $jsonDestinationIP = $matches[1].Trim()
    }

    # Host IP fallback
    if ($jsonText -match '(?i)"hostip"\s*:\s*"([^"]+)"') {
        $jsonHostIP = $matches[1].Trim()
    }

    # -----------------------------
    # DETECTION TYPE
    # -----------------------------

    $detectionType = ""

    
    if ($text -match '(?i)Sophos') {
        $detectionType = "EDR Alert"
    }
    elseif ($text -cmatch 'XDR') {
        $detectionType = "SIEM Correlation"
    }

    # -----------------------------
    # GET CURRENT CHROME URL
    # -----------------------------

    $urlValue = ""

    $wshell = New-Object -ComObject WScript.Shell

    # Activate Chrome
    $wshell.AppActivate("Google Chrome") | Out-Null

    Start-Sleep -Milliseconds 150

    # Focus address bar
    [System.Windows.Forms.SendKeys]::SendWait("^l")

    Start-Sleep -Milliseconds 75

    # Copy current URL
    [System.Windows.Forms.SendKeys]::SendWait("^c")

    Start-Sleep -Milliseconds 100

    # Read URL from clipboard
    $urlValue = [System.Windows.Forms.Clipboard]::GetText()

    #Check if URL is working
    #[System.Windows.Forms.MessageBox]::Show("URL captured:`n$urlValue")
    
    # Read URL from clipboard
    $urlValue = [System.Windows.Forms.Clipboard]::GetText()

    # Remove the question mark and everything after it
    if ($urlValue -match '\?') {
        $urlValue = $urlValue.Split('?')[0]
    }


    
    # -----------------------------
    # STELLAR #
    # -----------------------------

    $stellar = ""

    if ($text -match '(?m)^\s*(\d+)\s*:') {
        $stellar = $matches[1]
    }


    # -----------------------------
    # ISSUE
    # -----------------------------

    $issue = ""

    if ($text -match '(?s)^\s*\d+\s*:\s*(.*?)\s*Run Analysis') {
        $issue = $matches[1].Trim()
    }


    # -----------------------------
    # WHO
    # -----------------------------

    $who = ""

    if ($text -match '(?ms)^\s*Who\s*\r?\n\s*(.+?)\s*$') {
        $who = $matches[1].Trim()
    }


    # -----------------------------
    # WHAT
    # -----------------------------

    $what = ""

    if ($text -match '(?ms)^\s*What\s*\r?\n\s*(.+?)\s*$') {
        $what = $matches[1].Trim()
    }


    # -----------------------------
    # WHEN
    # -----------------------------

    $when = ""

    if ($text -match '(?ms)^\s*When\s*\r?\n\s*(.+?)\s*$') {
        $when = $matches[1].Trim()
    }


    # -----------------------------
    # WHERE
    # -----------------------------

    $where = ""

    if ($text -match '(?ms)^\s*Where\s*\r?\n\s*(.+?)\s*$') {
        $where = $matches[1].Trim()
    }

    # If Where is blank and parser accidentally grabs "Severity",
    # use Idaho as the fallback location
    if ([string]::IsNullOrWhiteSpace($where) -or $where -ieq "Severity") {
        $where = "Idaho"
    }

    # -----------------------------
    # SEVERITY
    # -----------------------------

    $severity = ""

    if ($text -match '(?ms)^\s*Severity\s*\r?\n\s*(.+?)\s*$') {
        $severity = $matches[1].Trim()
    }


    # -----------------------------
    # TENANT NAME
    # -----------------------------
    <#
    $tenantName = ""

    if ($text -match '(?mi)^\s*Tenant:\s*(.+?)\s*$') {
        $tenantName = $matches[1].Trim()
    }
    #>



    # -----------------------------
    # CASE SCORE BREAKDOWN
    # -----------------------------

    $caseScoreBreakdown = ""

    if ($text -match '(?ms)(Case Score Breakdown\s*\r?\nObserved[^\r\n]*(?:\r?\nInvolved[^\r\n]*)*)') {

        $caseScoreBreakdown = $matches[1].Trim()

    }

    # -----------------------------
    # SOURCE HOST
    # First IP address in Case Score Breakdown
    # Supports IPv4 and IPv6
    # -----------------------------

    $sourceHost = ""

    if ($caseScoreBreakdown) {

        # IPv4
        $ipv4Pattern = '\b(?:\d{1,3}\.){3}\d{1,3}\b'

        # IPv6
        $ipv6Pattern = '(?i)\b(?:[0-9a-f]{1,4}:){2,7}[0-9a-f]{0,4}\b'

        $ipMatch = [regex]::Match(
            $caseScoreBreakdown,
            "$ipv4Pattern|$ipv6Pattern"
        )

        if ($ipMatch.Success) {
            $sourceHost = $ipMatch.Value
        }
    }

    # Prefer JSON srcip when available
    if ($jsonSourceIP) {
        $sourceHost = $jsonSourceIP
    }
    # If no srcip, use hostip
    elseif ($jsonHostIP) {
        $sourceHost = $jsonHostIP
    }

    # -----------------------------
    # DESTINATION HOST
    # -----------------------------

    $destinationHost = ""

    if ($jsonDestinationIP) {
        $destinationHost = $jsonDestinationIP
    }
    elseif ($sourceHost) {
        $destinationHost = $sourceHost
    }

    # ==========================================
    # DESCRIPTION / SUPPORTING NOTES TEMPLATE
    # ==========================================

    $notesTemplate = "Summary:`r`n`r`n" +
                    "SIP: $sourceHost `r`n" +
                    "DIP: $destinationHost `r`n`r`n" +
                    "$caseScoreBreakdown`r`n`r`n" +
                    "Recommendation:`r`n"

    $descriptionText = $notesTemplate
    $supportingNotesText = $notesTemplate

    # -----------------------------
    # KILL CHAIN STAGE + MITRE TACTIC
    # Uses the FIRST Associated Alert
    # -----------------------------

    $killChainStage = ""
    $mitreTactic = ""

    if ($text -match '(?ms)Associated Alerts.*?\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s*\r?\n\s*([^\r\n]+)\s*\r?\n\s*([^\r\n]+)') {

        $killChainStage = $matches[1].Trim()
        $mitreTactic = $matches[2].Trim()

    }
    if ($mitreTactic -eq "Command and Control") {
        $mitreTactic = "Command & Control"
    } 


    $validMitreTactics = @(
        "Initial Access",
        "Execution",
        "Persistence",
        "Privilege Escalation",
        "Defense Evasion",
        "Credential Access",
        "Lateral Movement",
        "Command & Control"
    )


    # default MITRE Tactic to Initial Access
    if ($killChainStage -ieq "Initial Attempts" -and $mitreTactic -notin $validMitreTactics) {
        $mitreTactic = "Initial Access"
    }





    
    # ==========================================
    # BUILD AIRTABLE PREFILL URL
    # ==========================================

    $parameters = @()

    # Prefill selected analysts
    if ($analystChecklist.CheckedItems.Count -gt 0) {

        $selectedAnalysts = @(
            $analystChecklist.CheckedItems |
            ForEach-Object { $_.ToString() }
        )

        $analystValue = $selectedAnalysts -join ","

        $parameters += "prefill_Analyst=$([System.Web.HttpUtility]::UrlEncode($analystValue))"
    }

    if ($stellar) {
        $parameters += "prefill_Stellar%20%23=$([uri]::EscapeDataString($stellar))"
    }

    if ($issue) {
        $issueEncoded = [System.Web.HttpUtility]::UrlEncode($issue)
        $parameters += "prefill_Issue=$issueEncoded"
    }

    if ($who) {
        $parameters += "prefill_Who=$([uri]::EscapeDataString($who))"
    }

    if ($what) {
        $parameters += "prefill_What=$([uri]::EscapeDataString($what))"
    }

    if ($when) {
        $parameters += "prefill_When=$([uri]::EscapeDataString($when))"
    }

    if ($where) {
        $parameters += "prefill_Where=$([uri]::EscapeDataString($where))"
    }

    if ($urlValue) {
        $parameters += "prefill_fldxoVWezvmqx3QV2=$([System.Web.HttpUtility]::UrlEncode($urlValue))"
    }

    if ($severity) {
        $parameters += "prefill_Stellar+Severity=$([uri]::EscapeDataString($severity))"
    }

    # Always set Escalation Required? to Mentor Review
    $parameters += "prefill_Escalation+Required%3F=Mentor+Review"

   <# if ($tenantName) {
        $parameters += "prefill_Tenant+Name=$([uri]::EscapeDataString($tenantName))"
    } #>


   # Description only if Case Score Breakdown exists
    if ($caseScoreBreakdown) {
        $parameters += "prefill_Description=$([System.Web.HttpUtility]::UrlEncode($descriptionText))"
    }

    # Supporting Notes should ALWAYS be populated
    $parameters += "prefill_Supporting+Notes=$([System.Web.HttpUtility]::UrlEncode($supportingNotesText))"
    
    if ($killChainStage) {
        $parameters += "prefill_Kill+Chain+Stage=$([System.Web.HttpUtility]::UrlEncode($killChainStage))"
    }

    if ($mitreTactic) {
        $parameters += "prefill_MITRE+Tactic=$([System.Web.HttpUtility]::UrlEncode($mitreTactic))"
    }

    if ($sourceHost) {
        $parameters += "prefill_Source+Host%3A=$([System.Web.HttpUtility]::UrlEncode($sourceHost))"
    }

    if ($destinationHost) {
        $parameters += "prefill_Destination+Host%3A=$([System.Web.HttpUtility]::UrlEncode($destinationHost))"
    }

    if ($detectionType) {
        $parameters += "prefill_Detection+Type=$([System.Web.HttpUtility]::UrlEncode($detectionType))"
    }

    # Try to include JSON only if it will not make the URL too long
    if ($jsonText) {

        $jsonParameter = "prefill_JSON+of+the+event%3A=$([System.Web.HttpUtility]::UrlEncode($jsonText))"

        # Build a test URL with JSON included
        $testParameters = $parameters + $jsonParameter
        $testUrl = $airtableFormUrl + "?" + ($testParameters -join "&")

        $testUrlByteLength = [System.Text.Encoding]::UTF8.GetByteCount($testUrl)

        # Keep JSON only if the complete URL stays below the safe limit
        if ($testUrlByteLength -le 7800) {
            $parameters += $jsonParameter
        }
    }


    

   

    


    $finalUrl = $airtableFormUrl + "?" + ($parameters -join "&")

    <#
    $urlByteLength = [System.Text.Encoding]::UTF8.GetByteCount($finalUrl)

    [System.Windows.Forms.MessageBox]::Show(
        "URL length: $urlByteLength bytes"
    )
    #>

    # ==========================================
    # OPEN IN CHROME
    # ==========================================

    Start-Process "chrome.exe" $finalUrl

})

# ==========================================
# KEYBOARD SHORTCUTS
# ==========================================

$textBox.Add_KeyDown({

    # Backspace clears the entire textbox
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Back) {
        $textBox.Clear()
        $_.SuppressKeyPress = $true
    }

    # Enter runs the Fill Airtable button
    elseif ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $button.PerformClick()
        $_.SuppressKeyPress = $true
    }

})

# ==========================================
# SAVE ANALYST SETTINGS WHEN WINDOW CLOSES
# ==========================================

$form.Add_FormClosing({

    $settings = [ordered]@{}

    foreach ($analyst in $analysts) {

        $settings[$analyst] = (
            $analystChecklist.CheckedItems -contains $analyst
        )
    }

    $settings |
        ConvertTo-Json |
        Set-Content $analystSettingsFile
})
# ==========================================
# SHOW WINDOW
# ==========================================

$form.ShowDialog()
