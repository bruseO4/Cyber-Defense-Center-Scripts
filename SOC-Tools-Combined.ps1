# SOC Tools - Combined PowerShell GUI
# Home, IP Lookup, and Stellar to Airtable in one window.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

# Focus Airtable's rich-text JSON editor by its accessibility name, paste the
# complete JSON, then return to the beginning of Description. If Airtable has
# not finished loading, retry until the timeout expires. No alert is shown if
# focus is unavailable; the JSON remains on the clipboard for a normal Ctrl+V.
function Set-AirtableJsonField {
    param(
        [Parameter(Mandatory = $true)]
        [string]$JsonText,

        [int]$TimeoutSeconds = 20
    )

    if ([string]::IsNullOrWhiteSpace($JsonText)) {
        return $false
    }

    try {
        [System.Windows.Forms.Clipboard]::SetText($JsonText)
    }
    catch {
        return $false
    }

    $airtableShell = New-Object -ComObject WScript.Shell
    $airtableDeadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)

    while ([DateTime]::UtcNow -lt $airtableDeadline) {
        $airtableActivated = $airtableShell.AppActivate("Interface Form - Airtable")

        if (-not $airtableActivated) {
            $airtableActivated = $airtableShell.AppActivate("Google Chrome")
        }

        if ($airtableActivated) {
            Start-Sleep -Milliseconds 350

            try {
                # Start at Chrome's focused element, then walk up to its window.
                $airtableSearchRoot = [System.Windows.Automation.AutomationElement]::FocusedElement
                $airtableTreeWalker = [System.Windows.Automation.TreeWalker]::RawViewWalker

                while ($null -ne $airtableSearchRoot) {
                    $airtableParent = $airtableTreeWalker.GetParent($airtableSearchRoot)

                    if ($null -eq $airtableParent -or $airtableParent.Current.ProcessId -eq 0) {
                        break
                    }

                    $airtableSearchRoot = $airtableParent
                }

                if ($null -ne $airtableSearchRoot) {
                    $airtableElements = $airtableSearchRoot.FindAll(
                        [System.Windows.Automation.TreeScope]::Descendants,
                        [System.Windows.Automation.Condition]::TrueCondition
                    )

                    for ($airtableIndex = 0; $airtableIndex -lt $airtableElements.Count; $airtableIndex++) {
                        $airtableElement = $airtableElements.Item($airtableIndex)

                        try {
                            if (
                                $airtableElement.Current.Name -ieq "Json of the event:" -and
                                $airtableElement.Current.IsEnabled -and
                                $airtableElement.Current.IsKeyboardFocusable
                            ) {
                                $airtableElement.SetFocus()
                                Start-Sleep -Milliseconds 150
                                [System.Windows.Forms.SendKeys]::SendWait("^v")

                                # After JSON is pasted, return to Description,
                                # scroll it into view, and place the cursor at
                                # the start of its leading blank lines.
                                Start-Sleep -Milliseconds 250
                                $airtableDescriptionElements = $airtableSearchRoot.FindAll(
                                    [System.Windows.Automation.TreeScope]::Descendants,
                                    [System.Windows.Automation.Condition]::TrueCondition
                                )

                                for (
                                    $airtableDescriptionIndex = 0;
                                    $airtableDescriptionIndex -lt $airtableDescriptionElements.Count;
                                    $airtableDescriptionIndex++
                                ) {
                                    $airtableDescriptionElement = $airtableDescriptionElements.Item(
                                        $airtableDescriptionIndex
                                    )

                                    try {
                                        if (
                                            $airtableDescriptionElement.Current.Name -ieq "Description:" -and
                                            $airtableDescriptionElement.Current.IsEnabled -and
                                            $airtableDescriptionElement.Current.IsKeyboardFocusable
                                        ) {
                                            $airtableDescriptionElement.SetFocus()
                                            Start-Sleep -Milliseconds 150
                                            [System.Windows.Forms.SendKeys]::SendWait("^{HOME}")
                                            return $true
                                        }
                                    }
                                    catch {
                                        # Airtable can refresh Description while focus changes.
                                    }
                                }

                                return $true
                            }
                        }
                        catch {
                            # Airtable can replace an accessibility element while loading.
                        }
                    }
                }
            }
            catch {
                # Keep retrying while Airtable builds the form and accessibility tree.
            }
        }

        Start-Sleep -Milliseconds 350
    }

    return $false
}

# Find the first complete, valid JSON object or array in pasted Stellar text.
# Braces inside quoted values are ignored, while the original indentation,
# spaces, and line breaks are preserved for readability in Airtable.
function Get-JsonFromText {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    for ($jsonStart = 0; $jsonStart -lt $Text.Length; $jsonStart++) {
        $jsonOpeningCharacter = $Text[$jsonStart]

        if ($jsonOpeningCharacter -ne '{' -and $jsonOpeningCharacter -ne '[') {
            continue
        }

        $jsonDepth = 0
        $jsonInString = $false
        $jsonEscaped = $false

        for ($jsonIndex = $jsonStart; $jsonIndex -lt $Text.Length; $jsonIndex++) {
            $jsonCharacter = $Text[$jsonIndex]

            if ($jsonInString) {
                if ($jsonEscaped) {
                    $jsonEscaped = $false
                    continue
                }

                if ($jsonCharacter -eq '\') {
                    $jsonEscaped = $true
                    continue
                }

                if ($jsonCharacter -eq '"') {
                    $jsonInString = $false
                }

                continue
            }

            if ($jsonCharacter -eq '"') {
                $jsonInString = $true
                continue
            }

            if ($jsonCharacter -eq '{' -or $jsonCharacter -eq '[') {
                $jsonDepth++
                continue
            }

            if ($jsonCharacter -ne '}' -and $jsonCharacter -ne ']') {
                continue
            }

            $jsonDepth--

            if ($jsonDepth -ne 0) {
                continue
            }

            $jsonCandidate = $Text.Substring(
                $jsonStart,
                $jsonIndex - $jsonStart + 1
            )

            try {
                # Validate the isolated section before using it as the event JSON.
                $null = ConvertFrom-Json -InputObject $jsonCandidate -ErrorAction Stop
            }
            catch {
                # This balanced section was not JSON. Continue looking for the
                # next complete object or array in the pasted text.
                break
            }

            return $jsonCandidate
        }
    }

    return ""
}

# Recalculate and then lock the text bounds of controls that were originally
# AutoSize. Windows Forms can repaint those controls with a smaller cached box
# after Chrome/Airtable takes focus, which clips the ends of labels and the
# bottom line of the private-IP reference.
function Repair-SocToolsTextBounds {
    param(
        [object[]]$Controls,
        [System.Windows.Forms.Form]$Form
    )

    if ($null -eq $Form -or $null -eq $Controls) {
        return
    }

    $Form.SuspendLayout()

    try {
        foreach ($socControl in $Controls) {
            if ($null -eq $socControl -or $socControl.IsDisposed) {
                continue
            }

            # Preserve the last known-good bounds before asking Windows for a
            # fresh preferred size. A later DPI repaint may report a smaller
            # preferred size, but this control is never allowed to shrink.
            $socCurrentWidth = $socControl.Width
            $socCurrentHeight = $socControl.Height

            $socControl.AutoSize = $true
            $socPreferredSize = $socControl.PreferredSize
            $socControl.AutoSize = $false

            $socLockedWidth = [Math]::Max(
                $socCurrentWidth,
                $socPreferredSize.Width + 8
            )
            $socLockedHeight = [Math]::Max(
                $socCurrentHeight,
                $socPreferredSize.Height + 4
            )
            $socControl.Size = New-Object System.Drawing.Size(
                $socLockedWidth,
                $socLockedHeight
            )
        }
    }
    finally {
        $Form.ResumeLayout($true)
    }

    $Form.Invalidate($true)
    $Form.Update()
}

# ============================================================
# HOME PAGE WEBSITE SETTINGS
# ============================================================

# Stellar Cyber login page
$stellarCyberUrl = "https://blackswan.stellarcyber.cloud/login"

# Excel checksheet stored in SharePoint. The Airtable form used by the
# Stellar to Airtable tab remains configured separately below.
$excelChecksheetUrl = "https://uscyberdefensecenter.sharepoint.com/:x:/r/sites/shadowsoc/_layouts/15/Doc.aspx?sourcedoc=%7B6491D190-06AC-47AA-BC83-B28085C6331A%7D&file=Check%20Sheet.xlsx&action=default&mobileredirect=true&wdExp=TEAMS-TREATMENT&web=1&CID=519968ED-47BA-459E-886D-90EEF46278B6"

# Optional Home-page websites. Their checkboxes start unchecked.
$googleClassroomUrl = "https://classroom.google.com/u/0/w/ODYyODI0MjY0MDEy/t/all"
$shadowSocManualUrl = "https://uscyberdefensecenter.sharepoint.com/:w:/s/shadowsoc/IQCZurtWkRGIRamxkfitx_LjAWaN887dxBlfJXxm3lhw9Ic?isSPOFile=1&ovuser=8d281d1d-9c4d-4bf7-b16e-032d15de9f6c%2Cblj210001%40utdallas.edu&wdExp=TEAMS-TREATMENT&web=1&TeamsCID=e0de9577-dc04-4547-bb73-1383c1d7dc65&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNjA2MTExODIxNiJ9&linkOpenTime=1783952921482"

# Paste the complete "Join Microsoft Teams Meeting" link between the quotes.
$teamsMeetingUrl = "https://teams.microsoft.com/l/chat/19:meeting_MWI5ZTMwZDMtZjBjYi00ODg2LTlmMjYtOGUyZTUyY2NkYWE4@thread.v2/conversations?context=%7B%22contextType%22%3A%22chat%22%7D"

# Opens the Microsoft Shifts app in the Teams desktop client. You will still
# choose Time Clock and press Clock in or Clock out yourself.
$teamsShiftsUrl = "msteams://teams.microsoft.com/l/entity/42f6c1da-a241-483a-a3cc-4f5be9185951/shifts"

# Chrome executable location
$mainChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
if (-not (Test-Path $mainChromePath)) {
    $mainChromePath = "chrome.exe"
}

# ============================================================
# LOCAL CLOCK AND ACTIVITY TRACKER SETTINGS
# ============================================================

# Remembers the selected Student Program Tracker workbook.
$trackerSettingsFile = "$PSScriptRoot\SOC-Tracker-Settings.json"

# Remembers an unfinished clock-in if SOC Tools is closed or restarted.
$clockStateFile = "$PSScriptRoot\SOC-Clock-State.json"

function Get-TrackerWorkbookPath {
    if (Test-Path $trackerSettingsFile) {
        try {
            $savedTrackerSettings = Get-Content $trackerSettingsFile -Raw | ConvertFrom-Json
            if ($savedTrackerSettings.TrackerPath -and (Test-Path $savedTrackerSettings.TrackerPath)) {
                return $savedTrackerSettings.TrackerPath
            }
        }
        catch {
            # If the saved setting cannot be read, ask for the workbook again.
        }
    }

    $trackerDialog = New-Object System.Windows.Forms.OpenFileDialog
    $trackerDialog.Title = "Select Your Student Program Tracker"
    $trackerDialog.Filter = "Excel Workbooks (*.xlsx;*.xlsm;*.xlsb)|*.xlsx;*.xlsm;*.xlsb|All Files (*.*)|*.*"
    $trackerDialog.CheckFileExists = $true
    $trackerDialog.Multiselect = $false

    if ($trackerDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    $trackerPath = $trackerDialog.FileName

    [PSCustomObject]@{
        TrackerPath = $trackerPath
    } |
        ConvertTo-Json |
        Set-Content $trackerSettingsFile

    return $trackerPath
}

function Open-TrackerWorkbookForUpdate {
    param([string]$TrackerPath)

    $trackerExcel = $null
    $trackerWorkbook = $null
    $trackerWorksheet = $null

    try {
        $trackerExcel = New-Object -ComObject Excel.Application
        $trackerExcel.Visible = $false
        $trackerExcel.DisplayAlerts = $false

        $trackerWorkbook = $trackerExcel.Workbooks.Open($TrackerPath, 0, $false)

        if ($trackerWorkbook.ReadOnly) {
            throw "The tracker workbook is read-only. Close it in Excel and try again."
        }

        $trackerWorksheet = $trackerWorkbook.Worksheets.Item("Activity Tracker")

        return [PSCustomObject]@{
            Excel     = $trackerExcel
            Workbook  = $trackerWorkbook
            Worksheet = $trackerWorksheet
        }
    }
    catch {
        if ($null -ne $trackerWorkbook) {
            $trackerWorkbook.Close($false)
        }
        if ($null -ne $trackerExcel) {
            $trackerExcel.Quit()
        }

        if ($null -ne $trackerWorksheet) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($trackerWorksheet)
        }
        if ($null -ne $trackerWorkbook) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($trackerWorkbook)
        }
        if ($null -ne $trackerExcel) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($trackerExcel)
        }

        throw
    }
}

function Close-TrackerWorkbook {
    param(
        $TrackerSession,
        [bool]$SaveChanges
    )

    if ($null -eq $TrackerSession) {
        return
    }

    try {
        if ($SaveChanges) {
            $TrackerSession.Workbook.Save()
        }
        $TrackerSession.Workbook.Close($false)
        $TrackerSession.Excel.Quit()
    }
    finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($TrackerSession.Worksheet)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($TrackerSession.Workbook)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($TrackerSession.Excel)
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Start-SocClock {
    if (Test-Path $clockStateFile) {
        try {
            $existingClockState = Get-Content $clockStateFile -Raw | ConvertFrom-Json
            if ($existingClockState.Active) {
                [void][System.Windows.Forms.MessageBox]::Show(
                    "You are already clocked in from $($existingClockState.ClockInDisplay).",
                    "Already Clocked In"
                )
                return $null
            }
        }
        catch {
            [void][System.Windows.Forms.MessageBox]::Show(
                "The saved clock state could not be read. Delete SOC-Clock-State.json and try again.",
                "Clock State Error"
            )
            return $null
        }
    }

    $trackerPath = Get-TrackerWorkbookPath
    if ([string]::IsNullOrWhiteSpace($trackerPath)) {
        return $null
    }

    $clockInTime = Get-Date
    $daysSinceMonday = (([int]$clockInTime.DayOfWeek + 6) % 7)
    $weekStart = $clockInTime.Date.AddDays(-$daysSinceMonday)
    $trackerSession = $null
    try {
        $trackerSession = Open-TrackerWorkbookForUpdate -TrackerPath $trackerPath

        # Row 4 is the example row. Real activity starts on row 5.
        $trackerRow = 5
        while ($trackerRow -le 1000 -and -not [string]::IsNullOrWhiteSpace(
            [string]$trackerSession.Worksheet.Cells.Item($trackerRow, 9).Text
        )) {
            $trackerRow++
        }

        if ($trackerRow -gt 1000) {
            throw "No empty Activity Tracker rows were found."
        }

        # H = Week Start, I = Date, J = Start Time, K = End Time, M = Hours
        $trackerSession.Worksheet.Cells.Item($trackerRow, 8).Value2 = $weekStart.ToOADate()
        $trackerSession.Worksheet.Cells.Item($trackerRow, 9).Value2 = $clockInTime.Date.ToOADate()
        $trackerSession.Worksheet.Cells.Item($trackerRow, 10).Value2 = $clockInTime.TimeOfDay.TotalDays
        [void]$trackerSession.Worksheet.Cells.Item($trackerRow, 11).ClearContents()
        $trackerSession.Worksheet.Cells.Item($trackerRow, 13).Formula = "=ROUND((K$trackerRow-J$trackerRow)*24,2)"

        Close-TrackerWorkbook -TrackerSession $trackerSession -SaveChanges $true
        $trackerSession = $null
        $clockInDisplay = $clockInTime.ToString("MM/dd/yyyy h:mm:ss tt")

        [PSCustomObject]@{
            Active         = $true
            TrackerPath    = $trackerPath
            Worksheet      = "Activity Tracker"
            Row            = $trackerRow
            ClockIn        = $clockInTime.ToString("o")
            ClockInDisplay = $clockInDisplay
        } |
            ConvertTo-Json |
            Set-Content $clockStateFile

        return [PSCustomObject]@{
            Time    = $clockInTime
            Row     = $trackerRow
            Display = $clockInDisplay
        }
    }
    catch {
        if ($null -ne $trackerSession) {
            Close-TrackerWorkbook -TrackerSession $trackerSession -SaveChanges $false
        }

        [void][System.Windows.Forms.MessageBox]::Show(
            "Clock in could not be saved:`n$($_.Exception.Message)",
            "Clock In Error"
        )
        return $null
    }
}

function Stop-SocClock {
    if (-not (Test-Path $clockStateFile)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "No active clock-in was found.",
            "Not Clocked In"
        )
        return $null
    }

    try {
        $clockState = Get-Content $clockStateFile -Raw | ConvertFrom-Json
    }
    catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "The saved clock state could not be read.",
            "Clock State Error"
        )
        return $null
    }

    if (-not $clockState.Active) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "No active clock-in was found.",
            "Not Clocked In"
        )
        return $null
    }

    $clockOutTime = Get-Date
    $trackerSession = $null

    try {
        $trackerSession = Open-TrackerWorkbookForUpdate -TrackerPath $clockState.TrackerPath
        $trackerRow = [int]$clockState.Row

        $trackerSession.Worksheet.Cells.Item($trackerRow, 11).Value2 = $clockOutTime.TimeOfDay.TotalDays
        $trackerSession.Worksheet.Cells.Item($trackerRow, 13).Formula = "=ROUND((K$trackerRow-J$trackerRow)*24,2)"

        Close-TrackerWorkbook -TrackerSession $trackerSession -SaveChanges $true
        $trackerSession = $null

        Remove-Item $clockStateFile -Force

        return [PSCustomObject]@{
            Time    = $clockOutTime
            Row     = $trackerRow
            Display = $clockOutTime.ToString("MM/dd/yyyy h:mm:ss tt")
        }
    }
    catch {
        if ($null -ne $trackerSession) {
            Close-TrackerWorkbook -TrackerSession $trackerSession -SaveChanges $false
        }

        [void][System.Windows.Forms.MessageBox]::Show(
            "Clock out could not be saved:`n$($_.Exception.Message)",
            "Clock Out Error"
        )
        return $null
    }
}

# ============================================================
# MAIN WINDOW AND TABS
# ============================================================

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "SOC Tools"
$mainForm.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
$mainForm.ClientSize = New-Object System.Drawing.Size(720, 665)
$mainForm.StartPosition = "CenterScreen"
$mainForm.MinimumSize = New-Object System.Drawing.Size(736, 704)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"

$homeTab = New-Object System.Windows.Forms.TabPage
$homeTab.Text = "Home"
$homeTab.BackColor = [System.Drawing.Color]::FromArgb(45, 48, 52)
$homeTab.ForeColor = [System.Drawing.Color]::FromArgb(235, 235, 235)

$ipTab = New-Object System.Windows.Forms.TabPage
$ipTab.Text = "IP Lookup"
$ipTab.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)

$stellarTab = New-Object System.Windows.Forms.TabPage
$stellarTab.Text = "Stellar to Airtable"
$stellarTab.BackColor = [System.Drawing.Color]::FromArgb(32, 34, 37)

[void]$tabControl.TabPages.Add($homeTab)
[void]$tabControl.TabPages.Add($ipTab)
[void]$tabControl.TabPages.Add($stellarTab)
$mainForm.Controls.Add($tabControl)

# ============================================================
# HOME TAB
# ============================================================

$homeTitle = New-Object System.Windows.Forms.Label
$homeTitle.Text = "SOC Starting Place"
$homeTitle.AutoSize = $true
$homeTitle.Location = New-Object System.Drawing.Point(35, 35)
$homeTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$homeTitle.ForeColor = [System.Drawing.Color]::FromArgb(90, 170, 235)
$homeTab.Controls.Add($homeTitle)

$homeInstructions = New-Object System.Windows.Forms.Label
$homeInstructions.Text = "Choose the websites you want to open, then click the button."
$homeInstructions.AutoSize = $true
$homeInstructions.Location = New-Object System.Drawing.Point(39, 85)
$homeInstructions.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$homeTab.Controls.Add($homeInstructions)

$homeStellarCyberCheckBox = New-Object System.Windows.Forms.CheckBox
$homeStellarCyberCheckBox.Text = "Stellar Cyber"
$homeStellarCyberCheckBox.Checked = $true
$homeStellarCyberCheckBox.AutoSize = $true
$homeStellarCyberCheckBox.Location = New-Object System.Drawing.Point(42, 135)
$homeTab.Controls.Add($homeStellarCyberCheckBox)

$homeExcelChecksheetCheckBox = New-Object System.Windows.Forms.CheckBox
$homeExcelChecksheetCheckBox.Text = "Excel Checksheet"
$homeExcelChecksheetCheckBox.Checked = $true
$homeExcelChecksheetCheckBox.AutoSize = $true
$homeExcelChecksheetCheckBox.Location = New-Object System.Drawing.Point(42, 175)
$homeTab.Controls.Add($homeExcelChecksheetCheckBox)

$homeGoogleClassroomCheckBox = New-Object System.Windows.Forms.CheckBox
$homeGoogleClassroomCheckBox.Text = "Google Classroom"
$homeGoogleClassroomCheckBox.Checked = $false
$homeGoogleClassroomCheckBox.AutoSize = $true
$homeGoogleClassroomCheckBox.Location = New-Object System.Drawing.Point(285, 135)
$homeTab.Controls.Add($homeGoogleClassroomCheckBox)

$homeShadowSocManualCheckBox = New-Object System.Windows.Forms.CheckBox
$homeShadowSocManualCheckBox.Text = "Shadow SOC Manual"
$homeShadowSocManualCheckBox.Checked = $false
$homeShadowSocManualCheckBox.AutoSize = $true
$homeShadowSocManualCheckBox.Location = New-Object System.Drawing.Point(285, 175)
$homeTab.Controls.Add($homeShadowSocManualCheckBox)

$homeOpenButton = New-Object System.Windows.Forms.Button
$homeOpenButton.Text = "Open Selected Websites"
$homeOpenButton.Size = New-Object System.Drawing.Size(220, 45)
$homeOpenButton.Location = New-Object System.Drawing.Point(40, 225)
$homeOpenButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$homeOpenButton.ForeColor = [System.Drawing.Color]::White
$homeOpenButton.FlatStyle = "Flat"
$homeOpenButton.FlatAppearance.BorderSize = 0
$homeOpenButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$homeOpenButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$homeTab.Controls.Add($homeOpenButton)

$homeOpenButton.Add_Click({
    $homeSelectedUrls = @()

    if ($homeStellarCyberCheckBox.Checked) {
        $homeSelectedUrls += $stellarCyberUrl
    }

    if ($homeGoogleClassroomCheckBox.Checked) {
        $homeSelectedUrls += $googleClassroomUrl
    }

    if ($homeExcelChecksheetCheckBox.Checked) {
        $homeSelectedUrls += $excelChecksheetUrl
    }

    if ($homeShadowSocManualCheckBox.Checked) {
        $homeSelectedUrls += $shadowSocManualUrl
    }

    if ($homeSelectedUrls.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Select at least one website.",
            "No Website Selected"
        )
        return
    }

    if ($homeSelectedUrls.Count -gt 0) {
        Start-Process -FilePath $mainChromePath -ArgumentList $homeSelectedUrls
    }
})

$homeMeetingButton = New-Object System.Windows.Forms.Button
$homeMeetingButton.Text = "Join Teams Meeting"
$homeMeetingButton.Size = New-Object System.Drawing.Size(220, 45)
$homeMeetingButton.Location = New-Object System.Drawing.Point(40, 300)
$homeMeetingButton.BackColor = [System.Drawing.Color]::FromArgb(88, 80, 190)
$homeMeetingButton.ForeColor = [System.Drawing.Color]::White
$homeMeetingButton.FlatStyle = "Flat"
$homeMeetingButton.FlatAppearance.BorderSize = 0
$homeMeetingButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$homeMeetingButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$homeTab.Controls.Add($homeMeetingButton)

$homeMeetingButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($teamsMeetingUrl)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Paste your complete Teams meeting link near the top of this script first.",
            "Teams Meeting Link Needed"
        )
        return
    }

    if ($teamsMeetingUrl -notmatch '^https://teams\.microsoft\.com/l/') {
        [System.Windows.Forms.MessageBox]::Show(
            "The meeting link should begin with https://teams.microsoft.com/l/",
            "Invalid Teams Meeting Link"
        )
        return
    }

    # The msteams protocol skips the browser selection page and opens the
    # meeting's pre-join screen in the Teams desktop application.
    $teamsAppMeetingUrl = $teamsMeetingUrl -replace '^https://', 'msteams://'
    Start-Process $teamsAppMeetingUrl
})

$homeClockInButton = New-Object System.Windows.Forms.Button
$homeClockInButton.Text = "Clock In"
$homeClockInButton.Size = New-Object System.Drawing.Size(160, 45)
$homeClockInButton.Location = New-Object System.Drawing.Point(290, 300)
$homeClockInButton.BackColor = [System.Drawing.Color]::FromArgb(45, 145, 85)
$homeClockInButton.ForeColor = [System.Drawing.Color]::White
$homeClockInButton.FlatStyle = "Flat"
$homeClockInButton.FlatAppearance.BorderSize = 0
$homeClockInButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$homeClockInButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$homeTab.Controls.Add($homeClockInButton)

$homeClockOutButton = New-Object System.Windows.Forms.Button
$homeClockOutButton.Text = "Clock Out"
$homeClockOutButton.Size = New-Object System.Drawing.Size(160, 45)
$homeClockOutButton.Location = New-Object System.Drawing.Point(470, 300)
$homeClockOutButton.BackColor = [System.Drawing.Color]::FromArgb(185, 70, 70)
$homeClockOutButton.ForeColor = [System.Drawing.Color]::White
$homeClockOutButton.FlatStyle = "Flat"
$homeClockOutButton.FlatAppearance.BorderSize = 0
$homeClockOutButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$homeClockOutButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$homeTab.Controls.Add($homeClockOutButton)

$homeClockStatusLabel = New-Object System.Windows.Forms.Label
$homeClockStatusLabel.AutoSize = $false
$homeClockStatusLabel.Size = New-Object System.Drawing.Size(590, 55)
$homeClockStatusLabel.Location = New-Object System.Drawing.Point(40, 365)
$homeClockStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
$homeClockStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$homeTab.Controls.Add($homeClockStatusLabel)

if (Test-Path $clockStateFile) {
    try {
        $startupClockState = Get-Content $clockStateFile -Raw | ConvertFrom-Json
        if ($startupClockState.Active) {
            $homeClockStatusLabel.Text = "Clocked in: $($startupClockState.ClockInDisplay)`nTracker row: $($startupClockState.Row)"
        }
        else {
            $homeClockStatusLabel.Text = "Not currently clocked in"
        }
    }
    catch {
        $homeClockStatusLabel.Text = "Clock status could not be read"
    }
}
else {
    $homeClockStatusLabel.Text = "Not currently clocked in"
}

$homeClockInButton.Add_Click({
    $clockInResult = Start-SocClock

    if ($null -ne $clockInResult) {
        $homeClockStatusLabel.Text = "Clocked in: $($clockInResult.Display)`nTracker row: $($clockInResult.Row)"

        try {
            Start-Process $teamsShiftsUrl
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Your tracker time was saved, but Microsoft Teams could not be opened.",
                "Teams Not Available"
            )
        }
    }
})

$homeClockOutButton.Add_Click({
    $clockOutResult = Stop-SocClock

    if ($null -ne $clockOutResult) {
        $homeClockStatusLabel.Text = "Clocked out: $($clockOutResult.Display)`nTracker row: $($clockOutResult.Row)"

        try {
            Start-Process $teamsShiftsUrl
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Your tracker time was saved, but Microsoft Teams could not be opened.",
                "Teams Not Available"
            )
        }
    }
})

# ============================================================
# IP LOOKUP TAB
# ============================================================

# Loads the Windows Forms library so PowerShell can create a GUI window
Add-Type -AssemblyName System.Windows.Forms

# Loads drawing tools used for window/button sizes and positions
Add-Type -AssemblyName System.Drawing


# ------------------------------------------------------------
# HISTORY FILE
# ------------------------------------------------------------

# Saves lookup history in the same folder as this PowerShell script
$ipHistoryFile = "$PSScriptRoot\IP-Lookup-History.txt"

# Saves which lookup websites are checked or unchecked
$ipSettingsFile = "$PSScriptRoot\IP-Lookup-Settings.json"

# Chrome executable location
$ipChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
# ------------------------------------------------------------
# LOOKUP WEBSITE DEFINITIONS
# ------------------------------------------------------------

# Each website has a display name and URL.
# {0} will later be replaced with the IP address.
$ipSites = @(
    
    [PSCustomObject]@{
        Name = "LevelBlue OTX"
        Url  = "https://otx.alienvault.com/indicator/ip/{0}"
    }

    [PSCustomObject]@{
        Name = "VirusTotal"
        Url  = "https://www.virustotal.com/gui/ip-address/{0}"
    }

    [PSCustomObject]@{
        Name = "IBM X-Force"
        Url  = "https://exchange.xforce.ibmcloud.com/ip/{0}"
    }

    [PSCustomObject]@{
        Name = "Cisco Talos"
        Url  = "https://www.talosintelligence.com/reputation_center/lookup?search={0}"
    }

    [PSCustomObject]@{
        Name = "Spamhaus"
        Url  = "https://check.spamhaus.org/results?query={0}"
    }

    [PSCustomObject]@{
        Name = "AbuseIPDB"
        Url  = "https://www.abuseipdb.com/check/{0}"
    }

    [PSCustomObject]@{
        Name = "Scamalytics"
        Url  = "https://scamalytics.com/ip/{0}"
    }

    [PSCustomObject]@{
        Name = "IPQualityScore"
        Url  = "https://www.ipqualityscore.com/ip-reputation-check/lookup/{0}"
    }

    [PSCustomObject]@{
        Name = "IPVoid"
        Url  = "https://www.ipvoid.com/scan/{0}/"
    }

    [PSCustomObject]@{
        Name = "Censys"
        Url  = "https://search.censys.io/hosts/{0}"
    }

    [PSCustomObject]@{
        Name = "IPinfo"
        Url  = "https://ipinfo.io/{0}"
    }
)

# Sorts the lookup websites alphabetically by name
$ipSites = $ipSites | Sort-Object Name

# ------------------------------------------------------------
# IP ADDRESS LABEL
# ------------------------------------------------------------

# Creates the "Enter IP Address:" label
$ipLabel = New-Object System.Windows.Forms.Label

# Text displayed above the search box
$ipLabel.Text = "Enter IP Address:"

# Automatically sizes the label to fit the text
$ipLabel.AutoSize = $true

# Sets the label position in the window
$ipLabel.Location = New-Object System.Drawing.Point(20,25)


# ------------------------------------------------------------
# IP ADDRESS SEARCH BOX
# ------------------------------------------------------------

# Creates the text box where the IP address is entered
$ipTextBox = New-Object System.Windows.Forms.TextBox

# Sets the size of the text box
$ipTextBox.Size = New-Object System.Drawing.Size(250,25)

# Sets the position of the text box
$ipTextBox.Location = New-Object System.Drawing.Point(20,50)


# ------------------------------------------------------------
# BACKSPACE = CLEAR ENTIRE SEARCH BOX
# ------------------------------------------------------------

# Detects when a key is pressed while typing in the IP box
$ipTextBox.Add_KeyDown({

    # Checks whether the Backspace key was pressed
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Back) {

        # Clears the entire IP address instead of deleting one character
        $ipTextBox.Clear()

        # Stops Windows from performing the normal Backspace action
        $_.SuppressKeyPress = $true
    }
})


# ------------------------------------------------------------
# RUN LOOKUP BUTTON
# ------------------------------------------------------------

# Creates the Run Lookup button
$ipButton = New-Object System.Windows.Forms.Button

# Text shown on the button
$ipButton.Text = "Run Lookup"

# Sets the button size
$ipButton.Size = New-Object System.Drawing.Size(100,30)

# Sets the button position
$ipButton.Location = New-Object System.Drawing.Point(285,48)


# ------------------------------------------------------------
# ENTER KEY = RUN LOOKUP
# ------------------------------------------------------------

# Makes pressing Enter perform the same action as clicking Run Lookup


# ------------------------------------------------------------
# PRIVATE IP ADDRESS REFERENCE
# ------------------------------------------------------------

# Creates the label showing the three private IPv4 ranges
$ipPrivateLabel = New-Object System.Windows.Forms.Label

# Text containing the private IPv4 ranges
$ipPrivateLabel.Text = "Private IPv4 ranges:`n10.0.0.0 - 10.255.255.255`n172.16.0.0 - 172.31.255.255`n192.168.0.0 - 192.168.255.255"

# Automatically sizes the label
$ipPrivateLabel.AutoSize = $true

# Sets the position of the private IP reference
$ipPrivateLabel.Location = New-Object System.Drawing.Point(20,90)

# ------------------------------------------------------------
# WEBSITE CHECKLIST
# ------------------------------------------------------------

# Label displayed above the website checklist
$ipSiteLabel = New-Object System.Windows.Forms.Label
$ipSiteLabel.Text = "Lookup Websites:"
$ipSiteLabel.AutoSize = $true
$ipSiteLabel.Location = New-Object System.Drawing.Point(20,160)


# Creates a checklist containing all lookup websites
$ipSiteChecklist = New-Object System.Windows.Forms.CheckedListBox
$ipSiteChecklist.Size = New-Object System.Drawing.Size(365,155)
$ipSiteChecklist.Location = New-Object System.Drawing.Point(20,180)

# One click checks or unchecks a site
$ipSiteChecklist.CheckOnClick = $true

# Removes the blue highlight after clicking a checklist item
$ipSiteChecklist.Add_MouseUp({
    $ipSiteChecklist.ClearSelected()
})

# Chrome cleanup instructions shown directly below the website checklist
$ipCloseTabsTip = New-Object System.Windows.Forms.Label
$ipCloseTabsTip.Text = 'Chrome tip: To close all lookup tabs, right-click the tab you want to keep, then choose "Close tabs to the right."'
$ipCloseTabsTip.AutoSize = $false
$ipCloseTabsTip.Size = New-Object System.Drawing.Size(650,40)
$ipCloseTabsTip.Location = New-Object System.Drawing.Point(20,520)
$ipCloseTabsTip.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)


# ------------------------------------------------------------
# LOAD SAVED WEBSITE SETTINGS
# ------------------------------------------------------------

# If settings have previously been saved, load them
$ipSavedSites = $null

if (Test-Path $ipSettingsFile) {
    try {
        $ipSavedSites = Get-Content $ipSettingsFile -Raw | ConvertFrom-Json
    }
    catch {
        $ipSavedSites = $null
    }
}


# Add every website to the checklist
foreach ($ipSite in $ipSites) {

    # Default to checked
    $ipShouldBeChecked = $true

    # If saved settings exist, use the previous setting
    if ($null -ne $ipSavedSites) {

        if ($ipSavedSites.PSObject.Properties.Name -contains $ipSite.Name) {
            $ipShouldBeChecked = [bool]$ipSavedSites.($ipSite.Name)
        }
    }

    # Add the website and set its checked state
    [void]$ipSiteChecklist.Items.Add(
        $ipSite.Name,
        $ipShouldBeChecked
    )
}

# ------------------------------------------------------------
# LOOKUP HISTORY LABEL
# ------------------------------------------------------------

# Creates the "Lookup History:" label
$ipHistoryLabel = New-Object System.Windows.Forms.Label

# Text displayed above the history box
$ipHistoryLabel.Text = "Lookup History:"

# Automatically sizes the label
$ipHistoryLabel.AutoSize = $true

# Sets the position of the history label
$ipHistoryLabel.Location = New-Object System.Drawing.Point(20,350)


# ------------------------------------------------------------
# LOOKUP HISTORY BOX
# ------------------------------------------------------------

# Creates the box that displays previous IP searches
$ipHistoryBox = New-Object System.Windows.Forms.ListBox

# Sets the history box size
$ipHistoryBox.Size = New-Object System.Drawing.Size(365,140)

# Sets the history box position
$ipHistoryBox.Location = New-Object System.Drawing.Point(20,370)


# ------------------------------------------------------------
# CLICK HISTORY ENTRY = PUT IP BACK IN SEARCH BOX
# ------------------------------------------------------------

# Runs whenever a history item is selected
$ipHistoryBox.Add_SelectedIndexChanged({

    # Makes sure something was actually selected
    if ($ipHistoryBox.SelectedItem) {

        # Gets the complete selected history line
        $ipSelectedEntry = $ipHistoryBox.SelectedItem.ToString()

        # Extracts the IP address from the end of the history entry
$ipHistoryIP = ($ipSelectedEntry -split '\s+')[-1]

# Checks whether the extracted value is a valid IPv4 or IPv6 address
$ipParsedIP = $null

if ([System.Net.IPAddress]::TryParse($ipHistoryIP, [ref]$ipParsedIP)) {

    # Places the IP address into the search box
    $ipTextBox.Text = $ipHistoryIP

    # Moves keyboard focus back to the search box
    $ipTextBox.Focus()

    # Places the cursor at the end of the IP address
    $ipTextBox.SelectionStart = $ipTextBox.Text.Length
}
    }
})


# ------------------------------------------------------------
# LOAD SAVED HISTORY
# ------------------------------------------------------------

# Checks whether the history file already exists
if (Test-Path $ipHistoryFile) {

    # Reads all previous searches from the history file
    $ipSavedHistory = Get-Content $ipHistoryFile

    # Adds each saved entry to the visible history box
    foreach ($ipEntry in $ipSavedHistory) {
        $ipHistoryBox.Items.Add($ipEntry)
    }
}

# ------------------------------------------------------------
# STYLING
# ------------------------------------------------------------

# Main window background
$ipTab.BackColor = [System.Drawing.Color]::FromArgb(245,248,252)

# Main text color
$ipLabel.ForeColor = [System.Drawing.Color]::FromArgb(35,45,55)

# Search box colors
$ipTextBox.BackColor = [System.Drawing.Color]::White
$ipTextBox.ForeColor = [System.Drawing.Color]::FromArgb(25,25,25)
$ipTextBox.BorderStyle = "FixedSingle"

# ------------------------------------------------------------
# RUN LOOKUP BUTTON
# ------------------------------------------------------------

# Makes the button flat instead of the default Windows style
$ipButton.FlatStyle = "Flat"

# Main button color
$ipButton.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)

# Button text color
$ipButton.ForeColor = [System.Drawing.Color]::White

# Removes the default border
$ipButton.FlatAppearance.BorderSize = 0

# Color when hovering over the button
$ipButton.FlatAppearance.MouseOverBackColor = `
    [System.Drawing.Color]::FromArgb(0,100,190)

# Color while clicking the button
$ipButton.FlatAppearance.MouseDownBackColor = `
    [System.Drawing.Color]::FromArgb(0,80,160)



# ------------------------------------------------------------
# PRIVATE IP REFERENCE
# ------------------------------------------------------------

# Gives the private IP section a warm orange color
$ipPrivateLabel.ForeColor = [System.Drawing.Color]::FromArgb(190,90,35)

# Makes the Chrome tab-cleanup instruction readable without overpowering the form
$ipCloseTabsTip.ForeColor = [System.Drawing.Color]::FromArgb(75,85,95)

# ------------------------------------------------------------
# LOOKUP WEBSITE SECTION
# ------------------------------------------------------------

# Blue section heading
$ipSiteLabel.ForeColor = [System.Drawing.Color]::FromArgb(35,95,160)

# Checklist colors
$ipSiteChecklist.BackColor = [System.Drawing.Color]::White
$ipSiteChecklist.ForeColor = [System.Drawing.Color]::FromArgb(35,35,35)
$ipSiteChecklist.BorderStyle = "FixedSingle"

# ------------------------------------------------------------
# HISTORY SECTION
# ------------------------------------------------------------

# Blue section heading
$ipHistoryLabel.ForeColor = [System.Drawing.Color]::FromArgb(35,95,160)

# History box colors
$ipHistoryBox.BackColor = [System.Drawing.Color]::White
$ipHistoryBox.ForeColor = [System.Drawing.Color]::FromArgb(35,35,35)
$ipHistoryBox.BorderStyle = "FixedSingle"

# ------------------------------------------------------------
# RUN LOOKUP ACTION
# ------------------------------------------------------------

# Everything inside this block happens when Run Lookup is clicked
# or when Enter is pressed
$ipButton.Add_Click({

    # Gets the IP address from the search box
    # Gets whatever the user entered
$ipInputText = $ipTextBox.Text.Trim()


# --------------------------------------------------------
# NORMALIZE IP INPUT
# --------------------------------------------------------

# Removes http:// or https:// if present
$ipInputText = $ipInputText -replace '^https?://', ''

# Removes anything after a forward slash
# Example: 8.8.8.8/test becomes 8.8.8.8
$ipInputText = $ipInputText.Split('/')[0]

# Removes a port number from IPv4 addresses
# Example: 8.8.8.8:443 becomes 8.8.8.8
if ($ipInputText -match '^(\d{1,3}(?:\.\d{1,3}){3}):\d+$') {
    $ipInputText = $matches[1]
}

# The cleaned result becomes the IP used by the rest of the program
$ipIp = $ipInputText.Trim()

# Updates the search box so you can see what was extracted
$ipTextBox.Text = $ipIp

    # Creates a variable PowerShell will use to validate the IP
    $ipValidIP = $null


    # --------------------------------------------------------
    # VALIDATE IP ADDRESS
    # --------------------------------------------------------

    # Checks whether the entered text is a valid IPv4 or IPv6 address
    if (-not [System.Net.IPAddress]::TryParse($ipIp, [ref]$ipValidIP)) {

        # Shows an error popup if the IP is invalid
        [System.Windows.Forms.MessageBox]::Show(
            "'$ipIp' is not a valid IP address.",
            "Invalid IP"
        )

        # Stops the lookup from continuing
        return
    }


    # --------------------------------------------------------
    # CREATE HISTORY ENTRY
    # --------------------------------------------------------

    # Gets the current date and time
    $ipTimestamp = Get-Date -Format "MM/dd/yyyy hh:mm:ss tt"

    # Combines the date/time and IP into one history line
    $ipHistoryEntry = "$ipTimestamp   $ipIp"

    # Adds the newest lookup to the top of the history box
    $ipHistoryBox.Items.Insert(0, $ipHistoryEntry)


    # --------------------------------------------------------
    # SAVE HISTORY TO FILE
    # --------------------------------------------------------

    # Creates an empty list to hold previous history
    $ipExistingHistory = @()

    # If the history file already exists, read it
    if (Test-Path $ipHistoryFile) {
        $ipExistingHistory = Get-Content $ipHistoryFile
    }

    # Writes the newest entry first, followed by all older entries
    @($ipHistoryEntry) + $ipExistingHistory | Set-Content $ipHistoryFile

    # --------------------------------------------------------
    # OPEN SELECTED LOOKUP WEBSITES IN CHROME
    # --------------------------------------------------------

    # Holds all selected lookup URLs
    $ipSelectedUrls = @()

    # Go through every website
    foreach ($ipSite in $ipSites) {

        # Only use websites that are checked
        if ($ipSiteChecklist.CheckedItems -contains $ipSite.Name) {

            # Insert the IP address into the site's URL
            $ipUrl = [string]::Format($ipSite.Url, $ipIp)

            # Add the finished URL to the list
            $ipSelectedUrls += $ipUrl
        }
    }

    # Only continue if at least one website is selected
    if ($ipSelectedUrls.Count -gt 0) {

        # Send all selected URLs to Chrome
        # If Chrome is already open, these should open as new tabs
        Start-Process `
            -FilePath $ipChromePath `
            -ArgumentList $ipSelectedUrls
    }
    


    # --------------------------------------------------------
    # RESET SEARCH BOX AFTER LOOKUP
    # --------------------------------------------------------

    # Clears the IP that was just searched
    $ipTextBox.Clear()

    # Places the cursor back in the search box
    $ipTextBox.Focus()
})


# ------------------------------------------------------------
# ADD CONTROLS TO WINDOW
# ------------------------------------------------------------

# Adds the IP label to the window
$ipTab.Controls.Add($ipLabel)

# Adds the IP search box
$ipTab.Controls.Add($ipTextBox)

# Adds the Run Lookup button
$ipTab.Controls.Add($ipButton)

# Adds the Lookup History label
$ipTab.Controls.Add($ipHistoryLabel)

# Adds the Lookup History box
$ipTab.Controls.Add($ipHistoryBox)

# Adds the Private IPv4 ranges reference
$ipTab.Controls.Add($ipPrivateLabel)

# Adds the Lookup Websites label
$ipTab.Controls.Add($ipSiteLabel)

# Adds the website checklist
$ipTab.Controls.Add($ipSiteChecklist)

# Adds the Chrome tab cleanup instructions
$ipTab.Controls.Add($ipCloseTabsTip)


# ------------------------------------------------------------
# SAVE WEBSITE SETTINGS WHEN WINDOW CLOSES
# ------------------------------------------------------------

$mainForm.Add_FormClosing({

    # Creates an object that will store each website's
    # checked or unchecked state
    $ipSettings = [ordered]@{}

    foreach ($ipSite in $ipSites) {

        # True = checked
        # False = unchecked
        $ipSettings[$ipSite.Name] = (
            $ipSiteChecklist.CheckedItems -contains $ipSite.Name
        )
    }

    # Saves the settings as JSON
    $ipSettings |
        ConvertTo-Json |
        Set-Content $ipSettingsFile
})

# ============================================================
# STELLAR TO AIRTABLE TAB
# ============================================================

# ==========================================
# AIRTABLE SETTINGS
# ==========================================

$stellarAirtableFormUrl = "https://airtable.com/appybOSI4sqIAk36T/pagFt8aEkGJbcE2oO/form"

# Saves which analysts are selected
$stellarAnalystSettingsFile = "$PSScriptRoot\Stellar-Airtable-Analysts.json"

# ==========================================
# ANALYST LIST
# ==========================================

$stellarAnalysts = @(
    "Bruce Jamail"
    "Ruth A Nolan"
    "Bashar Al Qaraghuli"
    "Travis Fletcher"
    "Shloka Jain"
    "Chi-Heng Chan"
)

# ==========================================
# INSTRUCTIONS
# ==========================================

$stellarLabel = New-Object System.Windows.Forms.Label
$stellarLabel.Text = "Paste the Stellar alert information below:"
$stellarLabel.Location = New-Object System.Drawing.Point(20, 20)
$stellarLabel.Size = New-Object System.Drawing.Size(500, 25)
$stellarTab.Controls.Add($stellarLabel) 

# ==========================================
# TEXT BOX
# ==========================================

$stellarTextBox = New-Object System.Windows.Forms.TextBox
$stellarTextBox.Location = New-Object System.Drawing.Point(20, 50)
$stellarTextBox.Size = New-Object System.Drawing.Size(640, 360)
$stellarTextBox.Multiline = $true
$stellarTextBox.ScrollBars = "Vertical"
$stellarTab.Controls.Add($stellarTextBox)

# ==========================================
# ANALYST CHECKLIST
# ==========================================

$stellarAnalystLabel = New-Object System.Windows.Forms.Label
$stellarAnalystLabel.Text = "Analysts on Shift:"
$stellarAnalystLabel.Location = New-Object System.Drawing.Point(20, 420)
$stellarAnalystLabel.Size = New-Object System.Drawing.Size(200, 25)
$stellarAnalystLabel.ForeColor = [System.Drawing.Color]::FromArgb(90, 160, 220)
$stellarTab.Controls.Add($stellarAnalystLabel)

$stellarAnalystChecklist = New-Object System.Windows.Forms.CheckedListBox
$stellarAnalystChecklist.Location = New-Object System.Drawing.Point(20, 445)

# Stretch across almost the full window
# Height is sized for about 3 analyst rows
$stellarAnalystChecklist.Size = New-Object System.Drawing.Size(640, 65)

# After 3 rows, continue into another column
$stellarAnalystChecklist.MultiColumn = $true
$stellarAnalystChecklist.ColumnWidth = 200

$stellarAnalystChecklist.CheckOnClick = $true
$stellarAnalystChecklist.BackColor = [System.Drawing.Color]::FromArgb(38, 42, 46)
$stellarAnalystChecklist.ForeColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
$stellarAnalystChecklist.BorderStyle = "FixedSingle"
$stellarTab.Controls.Add($stellarAnalystChecklist)

# Load saved analyst selections
$stellarSavedAnalysts = $null

if (Test-Path $stellarAnalystSettingsFile) {
    try {
        $stellarSavedAnalysts = Get-Content $stellarAnalystSettingsFile -Raw | ConvertFrom-Json
    }
    catch {
        $stellarSavedAnalysts = $null
    }
}

foreach ($stellarAnalyst in $stellarAnalysts) {

    # Default to unchecked
    $stellarShouldBeChecked = $false

    # Use saved setting if one exists
    if ($null -ne $stellarSavedAnalysts) {
        if ($stellarSavedAnalysts.PSObject.Properties.Name -contains $stellarAnalyst) {
            $stellarShouldBeChecked = [bool]$stellarSavedAnalysts.($stellarAnalyst)
        }
    }

    [void]$stellarAnalystChecklist.Items.Add(
        $stellarAnalyst,
        $stellarShouldBeChecked
    )
}


# ==========================================
# FILL AIRTABLE BUTTON
# ==========================================

$stellarButton = New-Object System.Windows.Forms.Button
$stellarButton.Text = "Fill Airtable"
$stellarButton.Location = New-Object System.Drawing.Point(250, 535)
$stellarButton.Size = New-Object System.Drawing.Size(180, 45)
$stellarTab.Controls.Add($stellarButton)

# ==========================================
# STYLING
# ==========================================

# Window
$stellarTab.BackColor = [System.Drawing.Color]::FromArgb(32, 34, 37)
$stellarTab.ForeColor = [System.Drawing.Color]::White
$stellarTab.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Instruction label
$stellarLabel.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$stellarLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)

# Main text box
$stellarTextBox.BackColor = [System.Drawing.Color]::FromArgb(45, 48, 52)
$stellarTextBox.ForeColor = [System.Drawing.Color]::White
$stellarTextBox.BorderStyle = "FixedSingle"
$stellarTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)

# Fill Airtable button
$stellarButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$stellarButton.ForeColor = [System.Drawing.Color]::White
$stellarButton.FlatStyle = "Flat"
$stellarButton.FlatAppearance.BorderSize = 0
$stellarButton.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$stellarButton.Cursor = [System.Windows.Forms.Cursors]::Hand 

$stellarButton.Add_MouseEnter({
    $stellarButton.BackColor = [System.Drawing.Color]::FromArgb(0, 140, 240)
})

$stellarButton.Add_MouseLeave({
    $stellarButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
})



# ==========================================
# BUTTON ACTION
# ==========================================

$stellarButton.Add_Click({

    $stellarText = $stellarTextBox.Text

    # -----------------------------
    # JSON DATA
    # -----------------------------

    $stellarJsonText = ""
    $stellarJsonSourceIP = ""
    $stellarJsonDestinationIP = ""
    $stellarJsonHostIP = ""

    # Isolate one valid JSON value without capturing Overview or another object.
    # Keep the JSON's original indentation and line breaks for readability.
    $stellarJsonText = Get-JsonFromText -Text $stellarText

    # Source IP
    if ($stellarJsonText -match '(?i)"srcip"\s*:\s*"([^"]+)"') {
        $stellarJsonSourceIP = $matches[1].Trim()
    }

    # Destination IP
    if ($stellarJsonText -match '(?i)"dstip"\s*:\s*"([^"]+)"') {
        $stellarJsonDestinationIP = $matches[1].Trim()
    }

    # Host IP fallback
    if ($stellarJsonText -match '(?i)"hostip"\s*:\s*"([^"]+)"') {
        $stellarJsonHostIP = $matches[1].Trim()
    }

    # -----------------------------
    # DETECTION TYPE
    # -----------------------------

    $stellarDetectionType = ""

    
    if ($stellarText -match '(?i)Sophos') {
        $stellarDetectionType = "EDR Alert"
    }
    elseif ($stellarText -cmatch 'XDR') {
        $stellarDetectionType = "SIEM Correlation"
    }

    # -----------------------------
    # GET CURRENT CHROME URL
    # -----------------------------

    $stellarUrlValue = ""

    $stellarWshell = New-Object -ComObject WScript.Shell

    # Activate Chrome
    $stellarWshell.AppActivate("Google Chrome") | Out-Null

    Start-Sleep -Milliseconds 150

    # Focus address bar
    [System.Windows.Forms.SendKeys]::SendWait("^l")

    Start-Sleep -Milliseconds 75

    # Copy current URL
    [System.Windows.Forms.SendKeys]::SendWait("^c")

    Start-Sleep -Milliseconds 100

    # Read URL from clipboard
    $stellarUrlValue = [System.Windows.Forms.Clipboard]::GetText()

    #Check if URL is working
    #[System.Windows.Forms.MessageBox]::Show("URL captured:`n$stellarUrlValue")
    
    # Read URL from clipboard
    $stellarUrlValue = [System.Windows.Forms.Clipboard]::GetText()

    # Remove the question mark and everything after it
    if ($stellarUrlValue -match '\?') {
        $stellarUrlValue = $stellarUrlValue.Split('?')[0]
    }


    
    # -----------------------------
    # STELLAR #
    # -----------------------------

    $stellarStellar = ""

    if ($stellarText -match '(?m)^\s*(\d+)\s*:') {
        $stellarStellar = $matches[1]
    }


    # -----------------------------
    # ISSUE
    # -----------------------------

    $stellarIssue = ""

    # The case-number line can appear after copied page headers. Search from
    # the start of any line instead of only the start of the entire paste.
    $stellarIssueMatch = [regex]::Match(
        $stellarText,
        '(?ms)^\s*\d+\s*:\s*(?<Issue>.*?)\s*^\s*Run Analysis\s*$'
    )

    if ($stellarIssueMatch.Success) {

        # Chrome can insert line breaks when an alert title wraps. Airtable's
        # Issue field should receive one clean value without extra whitespace.
        $stellarIssue = (
            $stellarIssueMatch.Groups["Issue"].Value -replace '\s+', ' '
        ).Trim()

    }

    # Fallback for Stellar layouts that copy Issue as a separate field label.
    if (-not $stellarIssue -and $stellarText -match '(?mi)^\s*Issue\s*:?[ \t]*(?:\r?\n[ \t]*)?([^\r\n]+?)\s*$') {
        $stellarIssue = ($matches[1] -replace '\s+', ' ').Trim()
    }


    # -----------------------------
    # WHO
    # -----------------------------

    $stellarWho = ""

    if ($stellarText -match '(?ms)^\s*Who\s*\r?\n\s*(.+?)\s*$') {
        $stellarWho = $matches[1].Trim()
    }


    # -----------------------------
    # WHAT
    # -----------------------------

    $stellarWhat = ""

    if ($stellarText -match '(?ms)^\s*What\s*\r?\n\s*(.+?)\s*$') {
        $stellarWhat = $matches[1].Trim()
    }


    # -----------------------------
    # WHEN
    # -----------------------------

    $stellarWhen = ""

    if ($stellarText -match '(?ms)^\s*When\s*\r?\n\s*(.+?)\s*$') {
        $stellarWhen = $matches[1].Trim()
    }


    # -----------------------------
    # WHERE
    # -----------------------------

    $stellarWhere = ""

    if ($stellarText -match '(?ms)^\s*Where\s*\r?\n\s*(.+?)\s*$') {
        $stellarWhere = $matches[1].Trim()
    }

    # If Where is blank and parser accidentally grabs "Severity",
    # use Idaho as the fallback location
    if ([string]::IsNullOrWhiteSpace($stellarWhere) -or $stellarWhere -ieq "Severity") {
        $stellarWhere = "Idaho"
    }

    # -----------------------------
    # SEVERITY
    # -----------------------------

    $stellarSeverity = ""

    if ($stellarText -match '(?ms)^\s*Severity\s*\r?\n\s*(.+?)\s*$') {
        $stellarSeverity = $matches[1].Trim()
    }


    # -----------------------------
    # TENANT NAME
    # -----------------------------

    $stellarTenantName = ""

    # Stellar copies this field in several layouts. Try the visible label,
    # common JSON keys, and finally the first standalone SSOC value.
    $stellarTenantPatterns = @(
        '(?mi)^[^\S\r\n]*Tenant[^\S\r\n]+Name[^\S\r\n]*:?[^\S\r\n]*(?:\r?\n[^\S\r\n]*)?([^\r\n]+?)[^\S\r\n]*$',
        '(?mi)^[^\S\r\n]*Tenant[^\S\r\n]*:?[^\S\r\n]*(?:\r?\n[^\S\r\n]*)?([^\r\n]+?)[^\S\r\n]*$',
        '(?i)"(?:tenant[_ ]?name|tenant)"\s*:\s*"([^"]+)"',
        '(?mi)^[^\S\r\n]*(SSOC[^\S\r\n]*[:\-][^\r\n]+?)[^\S\r\n]*$'
    )

    foreach ($stellarTenantPattern in $stellarTenantPatterns) {
        $stellarTenantMatch = [regex]::Match($stellarText, $stellarTenantPattern)

        if ($stellarTenantMatch.Success) {
            $stellarTenantName = ($stellarTenantMatch.Groups[1].Value -replace '\s+', ' ').Trim()
            break
        }
    }

    # Stellar may use "SSOC: city" while Airtable uses "SSOC-city".
    if ($stellarTenantName -match '(?i)^SSOC\s*[:\-]\s*(.+)$') {
        $stellarTenantName = "SSOC-$($matches[1].Trim())"
    }



    # -----------------------------
    # CASE SCORE BREAKDOWN
    # -----------------------------

    $stellarCaseScoreBreakdown = ""

    if ($stellarText -match '(?ms)(Case Score Breakdown\s*\r?\nObserved[^\r\n]*(?:\r?\nInvolved[^\r\n]*)*)') {

        $stellarCaseScoreBreakdown = $matches[1].Trim()

    }

    # -----------------------------
    # SOURCE HOST
    # First IP address in Case Score Breakdown
    # Supports IPv4 and IPv6
    # -----------------------------

    $stellarSourceHost = ""

    if ($stellarCaseScoreBreakdown) {

        # IPv4
        $stellarIpv4Pattern = '\b(?:\d{1,3}\.){3}\d{1,3}\b'

        # IPv6
        $stellarIpv6Pattern = '(?i)\b(?:[0-9a-f]{1,4}:){2,7}[0-9a-f]{0,4}\b'

        $stellarIpMatch = [regex]::Match(
            $stellarCaseScoreBreakdown,
            "$stellarIpv4Pattern|$stellarIpv6Pattern"
        )

        if ($stellarIpMatch.Success) {
            $stellarSourceHost = $stellarIpMatch.Value
        }
    }

    # Prefer JSON srcip when available
    if ($stellarJsonSourceIP) {
        $stellarSourceHost = $stellarJsonSourceIP
    }
    # If no srcip, use hostip
    elseif ($stellarJsonHostIP) {
        $stellarSourceHost = $stellarJsonHostIP
    }

    # -----------------------------
    # DESTINATION HOST
    # -----------------------------

    $stellarDestinationHost = ""

    if ($stellarJsonDestinationIP) {
        $stellarDestinationHost = $stellarJsonDestinationIP
    }
    elseif ($stellarSourceHost) {
        $stellarDestinationHost = $stellarSourceHost
    }

    # ==========================================
    # DESCRIPTION / SUPPORTING NOTES TEMPLATE
    # ==========================================

    $stellarNotesTemplate = "Summary:`r`n`r`n" +
                    "SIP: $stellarSourceHost `r`n" +
                    "DIP: $stellarDestinationHost `r`n`r`n" +
                    "$stellarCaseScoreBreakdown`r`n`r`n" +
                    "Recommendation:`r`n"

    $stellarSupportingNotesText = $stellarNotesTemplate
    # Description matches Supporting Notes, with three blank lines first.
    $stellarDescriptionText = "`r`n`r`n`r`n" + $stellarSupportingNotesText

    # -----------------------------
    # KILL CHAIN STAGE + MITRE TACTIC
    # Uses the FIRST Associated Alert
    # -----------------------------

    $stellarKillChainStage = ""
    $stellarMitreTactic = ""

    if ($stellarText -match '(?ms)Associated Alerts.*?\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s*\r?\n\s*([^\r\n]+)\s*\r?\n\s*([^\r\n]+)') {

        $stellarKillChainStage = $matches[1].Trim()
        $stellarMitreTactic = $matches[2].Trim()

    }
    if ($stellarMitreTactic -eq "Command and Control") {
        $stellarMitreTactic = "Command & Control"
    } 


    $stellarValidMitreTactics = @(
        "Reconnaissance",
        "Resource Development",
        "Initial Access",
        "Execution",
        "Persistence",
        "Privilege Escalation",
        "Stealth",
        "Defense Evasion",
        "Credential Access",
        "Discovery",
        "Lateral Movement",
        "Collection",
        "Command & Control",
        "Exfiltration",
        "Impact",
        "XDR NBA",
        "XDR UBA"
    )


    # default MITRE Tactic to Initial Access
    if ($stellarKillChainStage -ieq "Initial Attempts" -and $stellarMitreTactic -notin $stellarValidMitreTactics) {
        $stellarMitreTactic = "Initial Access"
    }





    
    # ==========================================
    # BUILD AIRTABLE PREFILL URL
    # ==========================================

    $stellarParameters = @()

    # Prefill selected analysts
    if ($stellarAnalystChecklist.CheckedItems.Count -gt 0) {

        $stellarSelectedAnalysts = @(
            $stellarAnalystChecklist.CheckedItems |
            ForEach-Object { $_.ToString() }
        )

        $stellarAnalystValue = $stellarSelectedAnalysts -join ","

        $stellarParameters += "prefill_Analyst=$([System.Web.HttpUtility]::UrlEncode($stellarAnalystValue))"
    }

    if ($stellarStellar) {
        $stellarParameters += "prefill_Stellar%20%23=$([uri]::EscapeDataString($stellarStellar))"
    }

    if ($stellarIssue) {
        $stellarIssueEncoded = [System.Web.HttpUtility]::UrlEncode($stellarIssue)
        $stellarParameters += "prefill_Issue%3A=$stellarIssueEncoded"
    }

    if ($stellarWho) {
        $stellarParameters += "prefill_Who=$([uri]::EscapeDataString($stellarWho))"
    }

    if ($stellarWhat) {
        $stellarParameters += "prefill_What=$([uri]::EscapeDataString($stellarWhat))"
    }

    if ($stellarWhen) {
        $stellarParameters += "prefill_When=$([uri]::EscapeDataString($stellarWhen))"
    }

    if ($stellarWhere) {
        $stellarParameters += "prefill_Where=$([uri]::EscapeDataString($stellarWhere))"
    }

    if ($stellarUrlValue) {
        $stellarParameters += "prefill_fldxoVWezvmqx3QV2=$([System.Web.HttpUtility]::UrlEncode($stellarUrlValue))"
    }

    if ($stellarSeverity) {
        $stellarParameters += "prefill_Stellar+Severity=$([uri]::EscapeDataString($stellarSeverity))"
    }

    # Always set Escalation Required? to Mentor Review
    $stellarParameters += "prefill_Escalation+Required%3F=Mentor+Review"

    if ($stellarTenantName) {
        # Airtable's exact linked-record field label is "Tenant Name:".
        $stellarParameters += "prefill_Tenant+Name%3A=$([System.Web.HttpUtility]::UrlEncode($stellarTenantName))"
    }


    # Airtable's exact form label is "Description:"; encode the colon in the key.
    # Populate it every time, just like Supporting Notes.
    $stellarParameters += "prefill_Description%3A=$([System.Web.HttpUtility]::UrlEncode($stellarDescriptionText))"

    # Supporting Notes should ALWAYS be populated
    $stellarParameters += "prefill_Supporting+Notes=$([System.Web.HttpUtility]::UrlEncode($stellarSupportingNotesText))"
    
    if ($stellarKillChainStage) {
        $stellarParameters += "prefill_Kill+Chain+Stage=$([System.Web.HttpUtility]::UrlEncode($stellarKillChainStage))"
    }

    if ($stellarMitreTactic) {
        $stellarParameters += "prefill_MITRE+Tactic=$([System.Web.HttpUtility]::UrlEncode($stellarMitreTactic))"
    }

    if ($stellarSourceHost) {
        $stellarParameters += "prefill_Source+Host%3A=$([System.Web.HttpUtility]::UrlEncode($stellarSourceHost))"
    }

    if ($stellarDestinationHost) {
        $stellarParameters += "prefill_Destination+Host%3A=$([System.Web.HttpUtility]::UrlEncode($stellarDestinationHost))"
    }

    if ($stellarDetectionType) {
        $stellarParameters += "prefill_Detection+Type=$([System.Web.HttpUtility]::UrlEncode($stellarDetectionType))"
    }

    <# OLD JSON URL-PREFILL METHOD - DISABLED
    # Including JSON in the URL made long Stellar events exceed Airtable's
    # practical URL limit. The replacement pastes JSON after the form loads.
    if ($stellarJsonForAirtable) {

        $stellarJsonParameter = "prefill_JSON+of+the+event%3A=$([System.Web.HttpUtility]::UrlEncode($stellarJsonForAirtable))"

        # Build a test URL with JSON included
        $stellarTestParameters = $stellarParameters + $stellarJsonParameter
        $stellarTestUrl = $stellarAirtableFormUrl + "?" + ($stellarTestParameters -join "&")

        $stellarTestUrlByteLength = [System.Text.Encoding]::UTF8.GetByteCount($stellarTestUrl)

        # Leave extra safety room for Airtable and browser request handling.
        if ($stellarTestUrlByteLength -le 6500) {
            $stellarParameters += $stellarJsonParameter
        }
        else {
            $stellarJsonNeedsManualPaste = $true
        }
    }
    #>


    

   

    


    $stellarFinalUrl = $stellarAirtableFormUrl + "?" + ($stellarParameters -join "&")

    <#
    $stellarUrlByteLength = [System.Text.Encoding]::UTF8.GetByteCount($stellarFinalUrl)

    [System.Windows.Forms.MessageBox]::Show(
        "URL length: $stellarUrlByteLength bytes"
    )
    #>

    # ==========================================
    # OPEN IN CHROME
    # ==========================================

    Start-Process "chrome.exe" $stellarFinalUrl

    # Copy all JSON, wait for Airtable, focus the named rich-text editor, and
    # paste it. JSON never becomes part of the URL, so its length cannot crash
    # the form. If focus times out, the complete JSON remains on the clipboard.
    if ($stellarJsonText) {
        [void](Set-AirtableJsonField -JsonText $stellarJsonText -TimeoutSeconds 20)
    }

    # Chrome can return focus with stale AutoSize bounds. Refresh and lock the
    # full text sizes before the user returns to the Home or IP Lookup tabs.
    Repair-SocToolsTextBounds `
        -Controls $socToolsTextBoundsControls `
        -Form $mainForm

    <# OLD OVERSIZED-JSON ALERT - DISABLED
    if ($stellarJsonNeedsManualPaste -and $stellarJsonText) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($stellarJsonText)

            [System.Windows.Forms.MessageBox]::Show(
                "The complete JSON was too large for a safe Airtable link.`r`n`r`nThe form opened without JSON so it would not crash. The full JSON is now copied to your clipboard.`r`n`r`nClick the 'JSON of the event:' field and press Ctrl+V.",
                "Full JSON Copied"
            )
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "The JSON was too large for the Airtable link, and it could not be copied automatically.",
                "JSON Copy Error"
            )
        }
    }
    #>

})

# ==========================================
# KEYBOARD SHORTCUTS
# ==========================================

$stellarTextBox.Add_KeyDown({

    # Backspace clears the entire textbox
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Back) {
        $stellarTextBox.Clear()
        $_.SuppressKeyPress = $true
    }

    # Enter runs the Fill Airtable button
    elseif ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $stellarButton.PerformClick()
        $_.SuppressKeyPress = $true
    }

})

# ==========================================
# SAVE ANALYST SETTINGS WHEN WINDOW CLOSES
# ==========================================

$mainForm.Add_FormClosing({

    $stellarSettings = [ordered]@{}

    foreach ($stellarAnalyst in $stellarAnalysts) {

        $stellarSettings[$stellarAnalyst] = (
            $stellarAnalystChecklist.CheckedItems -contains $stellarAnalyst
        )
    }

    $stellarSettings |
        ConvertTo-Json |
        Set-Content $stellarAnalystSettingsFile
})

# ============================================================
# TAB BEHAVIOR AND DISPLAY
# ============================================================

$tabControl.Add_SelectedIndexChanged({
    if ($tabControl.SelectedTab -eq $homeTab) {
        $mainForm.AcceptButton = $homeOpenButton
    }
    elseif ($tabControl.SelectedTab -eq $ipTab) {
        $mainForm.AcceptButton = $ipButton
        $ipTextBox.Focus()
    }
    elseif ($tabControl.SelectedTab -eq $stellarTab) {
        $mainForm.AcceptButton = $stellarButton
        $stellarTextBox.Focus()
    }
})

$socToolsTextBoundsControls = @(
    $homeTitle
    $homeInstructions
    $homeStellarCyberCheckBox
    $homeExcelChecksheetCheckBox
    $homeGoogleClassroomCheckBox
    $homeShadowSocManualCheckBox
    $ipLabel
    $ipPrivateLabel
    $ipSiteLabel
    $ipHistoryLabel
)

# Refresh the affected text bounds whenever SOC Tools regains focus after
# Chrome/Airtable, including when the JSON automation times out.
$mainForm.Add_Activated({
    Repair-SocToolsTextBounds `
        -Controls $socToolsTextBoundsControls `
        -Form $mainForm
})

# Calculate the correct bounds once at startup and lock them before display.
Repair-SocToolsTextBounds `
    -Controls $socToolsTextBoundsControls `
    -Form $mainForm

$mainForm.AcceptButton = $homeOpenButton
[void]$mainForm.ShowDialog()
