# SOC Tools - Combined PowerShell GUI
# Home, IP Lookup, and Stellar to Airtable in one window.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# This function runs ONLY in the separate PowerShell worker below. Loading or
# using UI Automation in the GUI process can initialize WPF's DPI awareness
# after the form is already visible, changing its scale and checkbox rendering.
# Focus Airtable's rich-text JSON editor by its accessibility name and paste the
# complete JSON. Issue is already complete in the Airtable URL, so the worker
# focuses that field and leaves its cursor at the end for typing.
# If Airtable has not finished loading, retry until the timeout expires.
function Invoke-AirtableJsonFieldWorker {
    param(
        [Parameter(Mandatory = $true)]
        [string]$JsonText,

        [int]$TimeoutSeconds = 20
    )

    if ([string]::IsNullOrWhiteSpace($JsonText)) {
        return $false
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes

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

                                # Issue arrived as one complete URL-prefilled value.
                                # Focus it for the analyst and move only the
                                # caret; do not modify its rich-text contents.
                                Start-Sleep -Milliseconds 250
                                $airtableIssueElements = $airtableSearchRoot.FindAll(
                                    [System.Windows.Automation.TreeScope]::Descendants,
                                    [System.Windows.Automation.Condition]::TrueCondition
                                )

                                for (
                                    $airtableIssueIndex = 0;
                                    $airtableIssueIndex -lt $airtableIssueElements.Count;
                                    $airtableIssueIndex++
                                ) {
                                    $airtableIssueElement = $airtableIssueElements.Item(
                                        $airtableIssueIndex
                                    )

                                    try {
                                        if (
                                            $airtableIssueElement.Current.Name -ieq "Issue:" -and
                                            $airtableIssueElement.Current.IsEnabled -and
                                            $airtableIssueElement.Current.IsKeyboardFocusable
                                        ) {
                                            $airtableIssueElement.SetFocus()
                                            Start-Sleep -Milliseconds 150
                                            [System.Windows.Forms.SendKeys]::SendWait("^{END}")

                                            return $true
                                        }
                                    }
                                    catch {
                                        # Airtable can refresh Issue while focus changes.
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

# Keep UI Automation and its DPI changes outside the process that owns the
# SOC Tools window. A separate runspace/thread would still share process DPI.
# Send the existing worker and its inputs through standard input, so long JSON
# and Unicode do not become command-line arguments or temporary files.
function Set-AirtableJsonField {
    param(
        [Parameter(Mandatory = $true)]
        [string]$JsonText,

        [ValidateRange(1, 300)]
        [int]$TimeoutSeconds = 20
    )

    if ([string]::IsNullOrWhiteSpace($JsonText)) {
        return $false
    }

    $airtableWorkerProcess = $null
    $airtableWorkerStarted = $false

    try {
        # Preserve the existing clipboard fallback if automatic pasting fails.
        [System.Windows.Forms.Clipboard]::SetText($JsonText)

        $airtableWorkerRequest = @{
            WorkerScript = ${function:Invoke-AirtableJsonFieldWorker}.ToString()
            JsonText = $JsonText
            TimeoutSeconds = $TimeoutSeconds
        }
        $airtableRequestXml = [System.Management.Automation.PSSerializer]::Serialize(
            $airtableWorkerRequest
        )
        $airtableRequestBase64 = [Convert]::ToBase64String(
            [System.Text.Encoding]::Unicode.GetBytes($airtableRequestXml)
        )

        # Only this fixed bootstrap is on the command line. The alert is data,
        # never interpolated into executable PowerShell source.
        $airtableWorkerBootstrap = @'
$ErrorActionPreference = 'Stop'
try {
    $requestXml = [System.Text.Encoding]::Unicode.GetString(
        [Convert]::FromBase64String([Console]::In.ReadToEnd())
    )
    $request = [System.Management.Automation.PSSerializer]::Deserialize($requestXml)
    $worker = [scriptblock]::Create($request.WorkerScript)
    $pasted = & $worker -JsonText $request.JsonText -TimeoutSeconds $request.TimeoutSeconds
    if ($pasted -eq $true) { exit 0 }
    exit 1
}
catch {
    exit 1
}
'@
        $airtableEncodedBootstrap = [Convert]::ToBase64String(
            [System.Text.Encoding]::Unicode.GetBytes($airtableWorkerBootstrap)
        )

        # Windows PowerShell includes the .NET Framework UI Automation stack.
        # Use its full path so this also works when the GUI runs from PS7/ISE.
        $airtablePowerShellPath = Join-Path -Path ([Environment]::GetFolderPath('System')) `
            -ChildPath 'WindowsPowerShell\v1.0\powershell.exe'
        $airtableStartInfo = New-Object System.Diagnostics.ProcessStartInfo
        $airtableStartInfo.FileName = $airtablePowerShellPath
        $airtableStartInfo.Arguments = '-NoLogo -NoProfile -NonInteractive -STA -WindowStyle Hidden -EncodedCommand ' + $airtableEncodedBootstrap
        $airtableStartInfo.UseShellExecute = $false
        $airtableStartInfo.CreateNoWindow = $true
        $airtableStartInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        $airtableStartInfo.RedirectStandardInput = $true

        $airtableWorkerProcess = New-Object System.Diagnostics.Process
        $airtableWorkerProcess.StartInfo = $airtableStartInfo
        $airtableWorkerStarted = $airtableWorkerProcess.Start()
        if (-not $airtableWorkerStarted) {
            return $false
        }

        $airtableWorkerProcess.StandardInput.Write($airtableRequestBase64)
        $airtableWorkerProcess.StandardInput.Close()

        # Allow startup/paste overhead, but stop a stalled accessibility call
        # so the hidden worker cannot linger and paste into a later session.
        $airtableWorkerWaitMs = ($TimeoutSeconds + 15) * 1000
        if (-not $airtableWorkerProcess.WaitForExit($airtableWorkerWaitMs)) {
            return $false
        }

        return ($airtableWorkerProcess.ExitCode -eq 0)
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $airtableWorkerProcess) {
            try {
                if ($airtableWorkerStarted -and -not $airtableWorkerProcess.HasExited) {
                    $airtableWorkerProcess.Kill()
                    [void]$airtableWorkerProcess.WaitForExit(2000)
                }
            }
            catch {
                # The worker may already have exited between the check and Kill.
            }
            finally {
                $airtableWorkerProcess.Dispose()
            }
        }
    }
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

# Stellar uses xdr_event.description for the dynamic paragraph displayed on an
# alert's Overview tab. Read every unique occurrence from the copied JSON so the
# paragraph can be added to Airtable without switching browser tabs.
function Get-StellarAlertDescription {
    param(
        [string]$JsonText
    )

    if ([string]::IsNullOrWhiteSpace($JsonText)) {
        return ""
    }

    try {
        $stellarJsonObject = ConvertFrom-Json -InputObject $JsonText -ErrorAction Stop
    }
    catch {
        return ""
    }

    $stellarObjectsToInspect = New-Object System.Collections.Queue
    $stellarObjectsToInspect.Enqueue($stellarJsonObject)
    $stellarDescriptionValues = New-Object System.Collections.Generic.List[string]

    while ($stellarObjectsToInspect.Count -gt 0) {
        $stellarCurrentObject = $stellarObjectsToInspect.Dequeue()

        if ($null -eq $stellarCurrentObject -or $stellarCurrentObject -is [string]) {
            continue
        }

        # JSON arrays can contain more than one alert record.
        if (
            $stellarCurrentObject -is [System.Collections.IEnumerable] -and
            $stellarCurrentObject -isnot [System.Collections.IDictionary] -and
            $stellarCurrentObject -isnot [System.Management.Automation.PSCustomObject]
        ) {
            foreach ($stellarArrayItem in $stellarCurrentObject) {
                if ($null -ne $stellarArrayItem) {
                    $stellarObjectsToInspect.Enqueue($stellarArrayItem)
                }
            }

            continue
        }

        $stellarCurrentProperties = @()

        if ($stellarCurrentObject -is [System.Collections.IDictionary]) {
            foreach ($stellarDictionaryKey in $stellarCurrentObject.Keys) {
                $stellarCurrentProperties += [pscustomobject]@{
                    Name  = [string]$stellarDictionaryKey
                    Value = $stellarCurrentObject[$stellarDictionaryKey]
                }
            }
        }
        else {
            $stellarCurrentProperties = @($stellarCurrentObject.PSObject.Properties)
        }

        foreach ($stellarCurrentProperty in $stellarCurrentProperties) {
            $stellarPropertyName = [string]$stellarCurrentProperty.Name
            $stellarPropertyValue = $stellarCurrentProperty.Value

            # Stellar commonly exports the full field name as one JSON key.
            if (
                $stellarPropertyName -ieq "xdr_event.description" -and
                $stellarPropertyValue -is [string]
            ) {
                $stellarDescriptionCandidate = (
                    $stellarPropertyValue -replace "`r?`n", "`r`n"
                ).Trim()

                if (
                    -not [string]::IsNullOrWhiteSpace($stellarDescriptionCandidate) -and
                    $stellarDescriptionValues -notcontains $stellarDescriptionCandidate
                ) {
                    [void]$stellarDescriptionValues.Add($stellarDescriptionCandidate)
                }
            }

            # Also support JSON where xdr_event is an object containing description.
            if ($stellarPropertyName -ieq "xdr_event" -and $null -ne $stellarPropertyValue) {
                $stellarNestedDescriptionProperty = @(
                    $stellarPropertyValue.PSObject.Properties |
                    Where-Object { $_.Name -ieq "description" }
                ) | Select-Object -First 1

                if (
                    $null -ne $stellarNestedDescriptionProperty -and
                    $stellarNestedDescriptionProperty.Value -is [string]
                ) {
                    $stellarDescriptionCandidate = (
                        $stellarNestedDescriptionProperty.Value -replace "`r?`n", "`r`n"
                    ).Trim()

                    if (
                        -not [string]::IsNullOrWhiteSpace($stellarDescriptionCandidate) -and
                        $stellarDescriptionValues -notcontains $stellarDescriptionCandidate
                    ) {
                        [void]$stellarDescriptionValues.Add($stellarDescriptionCandidate)
                    }
                }
            }

            if (
                $null -ne $stellarPropertyValue -and
                $stellarPropertyValue -isnot [string] -and
                (
                    $stellarPropertyValue -is [System.Collections.IDictionary] -or
                    $stellarPropertyValue -is [System.Management.Automation.PSCustomObject] -or
                    $stellarPropertyValue -is [System.Collections.IEnumerable]
                )
            ) {
                $stellarObjectsToInspect.Enqueue($stellarPropertyValue)
            }
        }
    }

    return ($stellarDescriptionValues -join "`r`n`r`n")
}

# Measure and lock the startup text bounds once. Do not remeasure controls
# after browser focus changes; the window keeps its original layout and scale.
function Initialize-SocToolsTextBounds {
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

            # Keep the designed bounds if they are already larger than the
            # preferred text size, and add a little space for the final glyph.
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

function Get-TrackerRoleValue {
    param([string]$Role)

    switch ($Role) {
        "Shadow" { return "Shadow" }
        "L1"     { return "L1 Analyst" }
        "L2"     { return "L2 Analyst" }
        "L3"     { return "L3 Analyst" }
        default  { return $null }
    }
}

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
    param([string]$Role)

    $trackerRole = Get-TrackerRoleValue -Role $Role
    if ([string]::IsNullOrWhiteSpace($trackerRole)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Select Shadow, L1, L2, or L3 before clocking in.",
            "Role Needed"
        )
        return $null
    }

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

        # H = Week Start, I = Date, J = Start Time, K = End Time,
        # M = Hours, N = Role
        $trackerSession.Worksheet.Cells.Item($trackerRow, 8).Value2 = $weekStart.ToOADate()
        $trackerSession.Worksheet.Cells.Item($trackerRow, 9).Value2 = $clockInTime.Date.ToOADate()
        $trackerSession.Worksheet.Cells.Item($trackerRow, 10).Value2 = $clockInTime.TimeOfDay.TotalDays
        [void]$trackerSession.Worksheet.Cells.Item($trackerRow, 11).ClearContents()
        $trackerSession.Worksheet.Cells.Item($trackerRow, 13).Formula = "=ROUND((K$trackerRow-J$trackerRow)*24,2)"
        # Excel COM can reject a string assigned through Value2 on some systems.
        # Value accepts the role as text without forcing a numeric cast.
        $trackerSession.Worksheet.Cells.Item($trackerRow, 14).Value = $trackerRole

        Close-TrackerWorkbook -TrackerSession $trackerSession -SaveChanges $true
        $trackerSession = $null
        $clockInDisplay = $clockInTime.ToString("MM/dd/yyyy h:mm:ss tt")

        [PSCustomObject]@{
            Active         = $true
            TrackerPath    = $trackerPath
            Worksheet      = "Activity Tracker"
            Row            = $trackerRow
            Role           = $trackerRole
            RoleDisplay    = $Role
            ClockIn        = $clockInTime.ToString("o")
            ClockInDisplay = $clockInDisplay
        } |
            ConvertTo-Json |
            Set-Content $clockStateFile

        return [PSCustomObject]@{
            Time    = $clockInTime
            Row     = $trackerRow
            Role    = $trackerRole
            RoleDisplay = $Role
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
    param([string]$SelectedRole)

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

    $trackerRole = [string]$clockState.Role
    $roleDisplay = [string]$clockState.RoleDisplay

    # Supports an unfinished clock-in created by an older script version.
    if ([string]::IsNullOrWhiteSpace($trackerRole)) {
        $trackerRole = Get-TrackerRoleValue -Role $SelectedRole
        $roleDisplay = $SelectedRole
    }

    if ([string]::IsNullOrWhiteSpace($trackerRole)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Select Shadow, L1, L2, or L3 before clocking out.",
            "Role Needed"
        )
        return $null
    }

    try {
        $trackerSession = Open-TrackerWorkbookForUpdate -TrackerPath $clockState.TrackerPath
        $trackerRow = [int]$clockState.Row

        $trackerSession.Worksheet.Cells.Item($trackerRow, 11).Value2 = $clockOutTime.TimeOfDay.TotalDays
        $trackerSession.Worksheet.Cells.Item($trackerRow, 13).Formula = "=ROUND((K$trackerRow-J$trackerRow)*24,2)"
        $trackerSession.Worksheet.Cells.Item($trackerRow, 14).Value = $trackerRole

        Close-TrackerWorkbook -TrackerSession $trackerSession -SaveChanges $true
        $trackerSession = $null

        Remove-Item $clockStateFile -Force

        return [PSCustomObject]@{
            Time    = $clockOutTime
            Row     = $trackerRow
            Role    = $trackerRole
            RoleDisplay = $roleDisplay
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

$caseIpTab = New-Object System.Windows.Forms.TabPage
$caseIpTab.Text = "Case IP Lookups"
$caseIpTab.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)

[void]$tabControl.TabPages.Add($homeTab)
[void]$tabControl.TabPages.Add($ipTab)
[void]$tabControl.TabPages.Add($stellarTab)
[void]$tabControl.TabPages.Add($caseIpTab)
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
$homeMeetingButton.Text = "Join Team Meeting"
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

$homeRoleLabel = New-Object System.Windows.Forms.Label
$homeRoleLabel.Text = "Role"
$homeRoleLabel.AutoSize = $true
$homeRoleLabel.Location = New-Object System.Drawing.Point(350, 360)
$homeRoleLabel.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$homeTab.Controls.Add($homeRoleLabel)

$homeRoleComboBox = New-Object System.Windows.Forms.ComboBox
$homeRoleComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$homeRoleComboBox.Size = New-Object System.Drawing.Size(120, 30)
$homeRoleComboBox.Location = New-Object System.Drawing.Point(395, 353)
$homeRoleComboBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
[void]$homeRoleComboBox.Items.Add("Shadow")
[void]$homeRoleComboBox.Items.Add("L1")
[void]$homeRoleComboBox.Items.Add("L2")
[void]$homeRoleComboBox.Items.Add("L3")
$homeRoleComboBox.SelectedIndex = 1
$homeTab.Controls.Add($homeRoleComboBox)

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
$homeClockStatusLabel.Location = New-Object System.Drawing.Point(40, 400)
$homeClockStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
$homeClockStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$homeTab.Controls.Add($homeClockStatusLabel)

if (Test-Path $clockStateFile) {
    try {
        $startupClockState = Get-Content $clockStateFile -Raw | ConvertFrom-Json
        if ($startupClockState.Active) {
            if (-not [string]::IsNullOrWhiteSpace([string]$startupClockState.RoleDisplay)) {
                $homeRoleComboBox.SelectedItem = $startupClockState.RoleDisplay.ToString()
                $startupRoleText = " | Role: $($startupClockState.RoleDisplay)"
            }
            else {
                $startupRoleText = ""
            }

            $homeClockStatusLabel.Text = "Clocked in: $($startupClockState.ClockInDisplay)$startupRoleText`nTracker row: $($startupClockState.Row)"
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
    $selectedRole = "$($homeRoleComboBox.SelectedItem)"
    $clockInResult = Start-SocClock -Role $selectedRole

    if ($null -ne $clockInResult) {
        $homeClockStatusLabel.Text = "Clocked in: $($clockInResult.Display) | Role: $($clockInResult.RoleDisplay)`nTracker row: $($clockInResult.Row)"

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
    $selectedRole = "$($homeRoleComboBox.SelectedItem)"
    $clockOutResult = Stop-SocClock -SelectedRole $selectedRole

    if ($null -ne $clockOutResult) {
        $homeClockStatusLabel.Text = "Clocked out: $($clockOutResult.Display) | Role: $($clockOutResult.RoleDisplay)`nTracker row: $($clockOutResult.Row)"

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

# Stores optional API keys beside this script. Blank or missing values fall
# back to Windows environment variables with the same names.
$ipApiKeysFile = "$PSScriptRoot\IP-Lookup-API-Keys.json"

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
# THREAT AND LOCATION SUMMARY HELPERS
# ------------------------------------------------------------

# Optional API keys are read from IP-Lookup-API-Keys.json first, then from
# Windows environment variables if the matching JSON value is blank or missing:
#   IPINFO_TOKEN, OTX_API_KEY, VIRUSTOTAL_API_KEY, ABUSEIPDB_API_KEY
function Get-IpLookupApiKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Test-Path -LiteralPath $ipApiKeysFile) {
        try {
            $apiKeySettings = Get-Content -LiteralPath $ipApiKeysFile -Raw |
                ConvertFrom-Json
            $apiKeyProperty = $apiKeySettings.PSObject.Properties[$Name]

            if (
                $null -ne $apiKeyProperty -and
                -not [string]::IsNullOrWhiteSpace([string]$apiKeyProperty.Value)
            ) {
                return ([string]$apiKeyProperty.Value).Trim()
            }
        }
        catch {
            # A missing or invalid JSON file does not stop the lookup. The
            # environment-variable fallback below is still attempted.
        }
    }

    $value = [Environment]::GetEnvironmentVariable(
        $Name,
        [EnvironmentVariableTarget]::Process
    )

    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable(
            $Name,
            [EnvironmentVariableTarget]::User
        )
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
        $value = [Environment]::GetEnvironmentVariable(
            $Name,
            [EnvironmentVariableTarget]::Machine
        )
    }

    if ([string]::IsNullOrWhiteSpace($value)) {
        return $null
    }

    return $value.Trim()
}

function Invoke-IpLookupJsonRequest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [hashtable]$Headers = @{}
    )

    # Older Windows PowerShell installations may not enable TLS 1.2 by default.
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor
        [Net.SecurityProtocolType]::Tls12

    return Invoke-RestMethod `
        -Uri $Uri `
        -Method Get `
        -Headers $Headers `
        -TimeoutSec 6 `
        -ErrorAction Stop
}

# Extract one canonical IPv4 or IPv6 address from common analyst input. This
# accepts raw addresses, URLs, a missing URL colon such as https/8.8.8.8/,
# trailing punctuation, IPv4 ports, and bracketed IPv6 addresses with ports.
function ConvertTo-NormalizedIpAddress {
    param(
        [string]$InputText
    )

    if ([string]::IsNullOrWhiteSpace($InputText)) {
        return $null
    }

    $candidate = $InputText.Trim()

    # Remove surrounding prose punctuation without damaging IPv6 colons.
    $candidate = $candidate -replace '^[\s''"<>\(\{]+', ''
    $candidate = $candidate -replace '[\s''"<>\)\},;]+$', ''

    # Accept http://, https://, and the common missing-colon form https/.
    $candidate = $candidate -replace '(?i)^(?:https?|hxxps?)\s*:?[\\/]+', ''

    # A URL containing IPv6 must normally enclose the address in brackets.
    if ($candidate -match '^\[(?<Address>[^\]]+)\](?::\d+)?(?:[/?#].*)?[.,;]?$') {
        $candidate = $matches['Address']
    }
    else {
        # Remove a URL path, query, or fragment. A raw IPv6 address does not
        # use any of these delimiters.
        $candidate = ($candidate -split '[/?#]', 2)[0]

        # Remove :port only from an IPv4-looking host. Colons inside a raw IPv6
        # address are preserved.
        if ($candidate -match '^(?<Address>\d{1,3}(?:\.\d{1,3}){3}):\d+$') {
            $candidate = $matches['Address']
        }
    }

    # Analysts often paste a sentence-ending period with an otherwise valid IP.
    $candidate = $candidate.Trim().TrimEnd(
        [char[]]@('.', ',', ';', ')', ']', '}', '>', '"', "'")
    )

    $parsedAddress = $null
    $looksLikeIpv4 = $candidate -match '^\d{1,3}(?:\.\d{1,3}){3}$'
    $looksLikeIpv6 = $candidate.Contains(':')

    if (
        ($looksLikeIpv4 -or $looksLikeIpv6) -and
        [System.Net.IPAddress]::TryParse($candidate, [ref]$parsedAddress)
    ) {
        if ($parsedAddress.IsIPv4MappedToIPv6) {
            $parsedAddress = $parsedAddress.MapToIPv4()
        }

        return $parsedAddress.ToString()
    }

    return $null
}

function Test-IpAddressIsPublic {
    param(
        [Parameter(Mandatory = $true)]
        [System.Net.IPAddress]$Address
    )

    if ($Address.IsIPv4MappedToIPv6) {
        $Address = $Address.MapToIPv4()
    }

    $bytes = $Address.GetAddressBytes()

    if ($Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
        # Private, loopback, link-local, carrier-grade NAT, documentation,
        # benchmarking, multicast, and reserved IPv4 ranges.
        if ($bytes[0] -eq 10) { return $false }
        if ($bytes[0] -eq 127) { return $false }
        if ($bytes[0] -eq 169 -and $bytes[1] -eq 254) { return $false }
        if ($bytes[0] -eq 172 -and $bytes[1] -ge 16 -and $bytes[1] -le 31) { return $false }
        if ($bytes[0] -eq 192 -and $bytes[1] -eq 168) { return $false }
        if ($bytes[0] -eq 100 -and $bytes[1] -ge 64 -and $bytes[1] -le 127) { return $false }
        if ($bytes[0] -eq 192 -and $bytes[1] -eq 0 -and ($bytes[2] -eq 0 -or $bytes[2] -eq 2)) { return $false }
        if ($bytes[0] -eq 198 -and ($bytes[1] -eq 18 -or $bytes[1] -eq 19)) { return $false }
        if ($bytes[0] -eq 198 -and $bytes[1] -eq 51 -and $bytes[2] -eq 100) { return $false }
        if ($bytes[0] -eq 203 -and $bytes[1] -eq 0 -and $bytes[2] -eq 113) { return $false }
        if ($bytes[0] -eq 0 -or $bytes[0] -ge 224) { return $false }

        return $true
    }

    if ($Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
        if ($Address.Equals([System.Net.IPAddress]::IPv6Any)) { return $false }
        if ($Address.Equals([System.Net.IPAddress]::IPv6Loopback)) { return $false }
        if ($Address.IsIPv6LinkLocal -or $Address.IsIPv6SiteLocal -or $Address.IsIPv6Multicast) { return $false }
        if (($bytes[0] -band 0xFE) -eq 0xFC) { return $false }

        # 2001:db8::/32 is reserved for documentation.
        if ($bytes[0] -eq 0x20 -and $bytes[1] -eq 0x01 -and $bytes[2] -eq 0x0D -and $bytes[3] -eq 0xB8) {
            return $false
        }

        return $true
    }

    return $false
}

function Get-IpAddressSourceLabel {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [int]$MatchIndex
    )

    $lineStart = $Text.LastIndexOf("`n", [Math]::Max(0, $MatchIndex - 1))
    if ($lineStart -lt 0) {
        $lineStart = 0
    }
    else {
        $lineStart++
    }

    $prefixLength = [Math]::Max(0, $MatchIndex - $lineStart)
    $linePrefix = $Text.Substring($lineStart, $prefixLength)
    # Preserve the exact JSON/property name immediately before the address.
    # Example:  "srcip2": "198.51.100.20"  becomes the label srcip2.
    $fieldMatch = [regex]::Match(
        $linePrefix,
        '(?i)(?<Field>[a-z_][a-z0-9_.-]*)\s*["'']?\s*[:=]\s*["'']?\s*$'
    )

    if ($fieldMatch.Success) {
        return $fieldMatch.Groups['Field'].Value
    }

    if ($linePrefix -match '(?i)(?:https?|hxxps?)\s*:?[\\/]+\s*$') {
        return "URL"
    }

    return "Alert text"
}

# Find every unique literal IPv4 and IPv6 address in copied Stellar text. Each
# result includes its source field(s), parsed value, and public/private status.
function Get-IpAddressesFromText {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    $candidateMatches = @()

    # Broad IPv4 candidates are validated by IPAddress.TryParse below.
    foreach ($match in [regex]::Matches(
        $Text,
        '(?<!\d)(?:\d{1,3}\.){3}\d{1,3}(?!\d)'
    )) {
        $candidateMatches += [PSCustomObject]@{
            Index = $match.Index
            Value = $match.Value
        }
    }

    # Collect colon-containing tokens, then let IPAddress.TryParse distinguish
    # IPv6 from timestamps, MAC addresses, and ordinary text.
    foreach ($match in [regex]::Matches(
        $Text,
        '(?i)(?<![0-9a-f:.%])[0-9a-f:.%]*:[0-9a-f:.%]+(?![0-9a-f:.%])'
    )) {
        $candidateMatches += [PSCustomObject]@{
            Index = $match.Index
            Value = $match.Value
        }
    }

    # Use native PowerShell collections here for full Windows PowerShell 5.1
    # compatibility. Generic HashSet constructor behavior differs between the
    # .NET Framework and newer PowerShell/.NET versions.
    $recordsByAddress = @{}
    $orderedAddressKeys = @()

    foreach ($candidateMatch in @($candidateMatches | Sort-Object Index)) {
        $normalizedAddress = ConvertTo-NormalizedIpAddress `
            -InputText $candidateMatch.Value

        $addressKey = [string]$normalizedAddress
        if (-not [string]::IsNullOrWhiteSpace($addressKey)) {
            $addressKey = $addressKey.ToLowerInvariant()
        }

        if ([string]::IsNullOrWhiteSpace($normalizedAddress)) {
            continue
        }

        $sourceLabel = Get-IpAddressSourceLabel `
            -Text $Text `
            -MatchIndex $candidateMatch.Index

        if ($recordsByAddress.ContainsKey($addressKey)) {
            $existingRecord = $recordsByAddress[$addressKey]
            if ($existingRecord.SourceLabels -notcontains $sourceLabel) {
                $existingRecord.SourceLabels = @($existingRecord.SourceLabels) + $sourceLabel
            }
            continue
        }

        $parsedAddress = $null
        if (-not [System.Net.IPAddress]::TryParse(
            $normalizedAddress,
            [ref]$parsedAddress
        )) {
            continue
        }

        $isPublicAddress = Test-IpAddressIsPublic -Address $parsedAddress

        $recordsByAddress[$addressKey] = [PSCustomObject]@{
            Address       = $normalizedAddress
            ParsedAddress = $parsedAddress
            IsPublic      = [bool]$isPublicAddress
            SourceLabels  = @($sourceLabel)
        }
        $orderedAddressKeys += $addressKey
    }

    $results = @(
        foreach ($addressKey in $orderedAddressKeys) {
            $record = $recordsByAddress[$addressKey]
            $sourceDisplay = @($record.SourceLabels) -join ", "

            [PSCustomObject]@{
                Address       = $record.Address
                ParsedAddress = $record.ParsedAddress
                IsPublic      = $record.IsPublic
                SourceLabels  = @($record.SourceLabels)
                Source        = $sourceDisplay
                Display       = "$($record.Address)  [$sourceDisplay]"
            }
        }
    )

    return $results
}

function Get-IpLookupCountryName {
    param([string]$CountryCode)

    if ([string]::IsNullOrWhiteSpace($CountryCode)) {
        return $null
    }

    try {
        return (New-Object System.Globalization.RegionInfo($CountryCode)).EnglishName
    }
    catch {
        return $CountryCode
    }
}

function ConvertFrom-IpLookupUnixTime {
    param($Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    try {
        $parsedTime = [DateTimeOffset]::FromUnixTimeSeconds([long]$Value)
        return $parsedTime.UtcDateTime.ToString(
            "yyyy-MM-dd HH:mm:ss 'UTC'"
        )
    }
    catch {
        return ([string]$Value).Trim()
    }
}

function ConvertTo-IpLookupUtcTime {
    param($Value)

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    try {
        $parsedTime = [DateTimeOffset]::Parse([string]$Value)
        return $parsedTime.UtcDateTime.ToString(
            "yyyy-MM-dd HH:mm:ss 'UTC'"
        )
    }
    catch {
        return ([string]$Value).Trim()
    }
}

function ConvertTo-IpLookupSingleLine {
    param(
        $Value,
        [int]$MaximumLength = 180
    )

    if ($null -eq $Value) {
        return $null
    }

    $text = (([string]$Value) -replace '\s+', ' ').Trim()

    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($MaximumLength -gt 3 -and $text.Length -gt $MaximumLength) {
        return $text.Substring(0, $MaximumLength - 3) + "..."
    }

    return $text
}

function Test-IpLookupValueReported {
    param(
        $InputObject,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    if ($null -eq $InputObject) {
        return $false
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]

    if ($null -eq $property -or $null -eq $property.Value) {
        return $false
    }

    if (
        $property.Value -is [string] -and
        [string]::IsNullOrWhiteSpace([string]$property.Value)
    ) {
        return $false
    }

    return $true
}

function ConvertTo-IpLookupYesNo {
    param($Value)

    if ([bool]$Value) {
        return "Yes"
    }

    return "No"
}

function Get-IpThreatLocationSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$IpAddress,

        [Parameter(Mandatory = $true)]
        [System.Net.IPAddress]$ParsedAddress,

        [switch]$ImportantOnly
    )

    $newLine = [Environment]::NewLine

    if (-not (Test-IpAddressIsPublic -Address $ParsedAddress)) {
        return @(
            "IP: $IpAddress"
            "OVERALL: PRIVATE OR RESERVED ADDRESS"
            ""
            "This address is not publicly routable, so it was not sent to external services."
            "Public geolocation and reputation results do not apply to it."
        ) -join $newLine
    }

    $encodedIp = [Uri]::EscapeDataString($IpAddress)
    $location = $null
    $network = $null
    $hostname = $null
    $providerResponses = 0
    $threatSourcesChecked = 0
    $threatSourcesWithAssessment = 0
    $hasPossibleThreatSignal = $false
    $hasHighThreatSignal = $false
    $socSignals = @()

    $ipInfoDetails = @("Status: Request unavailable.")
    $otxDetails = @("Status: Request unavailable.")
    $virusTotalDetails = @(
        "Status: Not checked. Add VIRUSTOTAL_API_KEY to the API key JSON file."
    )
    $abuseIpDbDetails = @(
        "Status: Not checked. Add ABUSEIPDB_API_KEY to the API key JSON file."
    )

    # IPinfo supplies the primary location and network-owner attribution.
    try {
        $ipInfoToken = Get-IpLookupApiKey -Name "IPINFO_TOKEN"
        $ipInfoUri = "https://ipinfo.io/$encodedIp/json"
        $ipInfoDataTier = "Legacy/Core"

        if (-not [string]::IsNullOrWhiteSpace($ipInfoToken)) {
            $ipInfoUri += "?token=$([Uri]::EscapeDataString($ipInfoToken))"
        }

        $ipInfo = $null

        try {
            $ipInfo = Invoke-IpLookupJsonRequest -Uri $ipInfoUri
        }
        catch {
            # New IPinfo accounts use the Lite endpoint. It returns country and
            # ASN/owner data, while legacy accounts can also return city/region.
            if (-not [string]::IsNullOrWhiteSpace($ipInfoToken)) {
                $ipInfoLiteUri = "https://api.ipinfo.io/lite/$encodedIp`?token=$([Uri]::EscapeDataString($ipInfoToken))"
                $ipInfo = Invoke-IpLookupJsonRequest -Uri $ipInfoLiteUri
                $ipInfoDataTier = "Lite"
            }
            else {
                throw
            }
        }

        $ipInfoCity = [string]$ipInfo.city
        $ipInfoRegion = [string]$ipInfo.region
        $ipInfoCountry = [string]$ipInfo.country
        $ipInfoCountryCode = [string]$ipInfo.country_code
        $ipInfoAsn = $null
        $ipInfoAsName = [string]$ipInfo.as_name
        $ipInfoAsDomain = [string]$ipInfo.as_domain
        $ipInfoAsType = $null
        $ipInfoNetworkRange = $null
        $ipInfoContinent = [string]$ipInfo.continent
        $ipInfoContinentCode = [string]$ipInfo.continent_code
        $ipInfoLatitude = $null
        $ipInfoLongitude = $null
        $ipInfoPostalCode = [string]$ipInfo.postal
        $ipInfoTimeZone = [string]$ipInfo.timezone

        if ($null -ne $ipInfo.asn) {
            if ($ipInfo.asn -is [string] -or $ipInfo.asn -is [ValueType]) {
                $ipInfoAsn = [string]$ipInfo.asn
            }
            else {
                $ipInfoAsn = [string]$ipInfo.asn.asn
                if ([string]::IsNullOrWhiteSpace($ipInfoAsName)) { $ipInfoAsName = [string]$ipInfo.asn.name }
                if ([string]::IsNullOrWhiteSpace($ipInfoAsDomain)) { $ipInfoAsDomain = [string]$ipInfo.asn.domain }
                $ipInfoAsType = [string]$ipInfo.asn.type
                $ipInfoNetworkRange = [string]$ipInfo.asn.route
            }
        }

        if ($null -ne $ipInfo.geo) {
            if ([string]::IsNullOrWhiteSpace($ipInfoCity)) { $ipInfoCity = [string]$ipInfo.geo.city }
            if ([string]::IsNullOrWhiteSpace($ipInfoRegion)) { $ipInfoRegion = [string]$ipInfo.geo.region }
            if ([string]::IsNullOrWhiteSpace($ipInfoCountry)) { $ipInfoCountry = [string]$ipInfo.geo.country }
            if ([string]::IsNullOrWhiteSpace($ipInfoCountryCode)) { $ipInfoCountryCode = [string]$ipInfo.geo.country_code }
            if ($null -ne $ipInfo.geo.continent -and $ipInfo.geo.continent -isnot [string]) {
                if ([string]::IsNullOrWhiteSpace($ipInfoContinent)) { $ipInfoContinent = [string]$ipInfo.geo.continent.name }
                if ([string]::IsNullOrWhiteSpace($ipInfoContinentCode)) { $ipInfoContinentCode = [string]$ipInfo.geo.continent.code }
            }
            elseif ([string]::IsNullOrWhiteSpace($ipInfoContinent)) {
                $ipInfoContinent = [string]$ipInfo.geo.continent
            }
            if ([string]::IsNullOrWhiteSpace($ipInfoContinentCode)) { $ipInfoContinentCode = [string]$ipInfo.geo.continent_code }
            if ([string]::IsNullOrWhiteSpace($ipInfoPostalCode)) { $ipInfoPostalCode = [string]$ipInfo.geo.postal_code }
            if ([string]::IsNullOrWhiteSpace($ipInfoTimeZone)) { $ipInfoTimeZone = [string]$ipInfo.geo.timezone }
            if (Test-IpLookupValueReported -InputObject $ipInfo.geo -PropertyName "latitude") { $ipInfoLatitude = $ipInfo.geo.latitude }
            if (Test-IpLookupValueReported -InputObject $ipInfo.geo -PropertyName "longitude") { $ipInfoLongitude = $ipInfo.geo.longitude }
        }

        if ($null -ne $ipInfo.as) {
            if ([string]::IsNullOrWhiteSpace($ipInfoAsn)) { $ipInfoAsn = [string]$ipInfo.as.asn }
            if ([string]::IsNullOrWhiteSpace($ipInfoAsName)) { $ipInfoAsName = [string]$ipInfo.as.name }
            if ([string]::IsNullOrWhiteSpace($ipInfoAsDomain)) { $ipInfoAsDomain = [string]$ipInfo.as.domain }
            if ([string]::IsNullOrWhiteSpace($ipInfoAsType)) { $ipInfoAsType = [string]$ipInfo.as.type }
            if ([string]::IsNullOrWhiteSpace($ipInfoNetworkRange)) { $ipInfoNetworkRange = [string]$ipInfo.as.route }
        }

        if (
            ($null -eq $ipInfoLatitude -or $null -eq $ipInfoLongitude) -and
            -not [string]::IsNullOrWhiteSpace([string]$ipInfo.loc) -and
            [string]$ipInfo.loc -match '^\s*([^,]+),\s*(.+)\s*$'
        ) {
            $ipInfoLatitude = $matches[1]
            $ipInfoLongitude = $matches[2]
        }

        # The legacy response stores a two-letter code in country. Newer
        # responses store the full country name and a separate country_code.
        if ($ipInfoCountry.Length -eq 2) {
            if ([string]::IsNullOrWhiteSpace($ipInfoCountryCode)) {
                $ipInfoCountryCode = $ipInfoCountry
            }

            $ipInfoCountry = Get-IpLookupCountryName -CountryCode $ipInfoCountry
        }
        elseif ([string]::IsNullOrWhiteSpace($ipInfoCountry)) {
            $ipInfoCountry = Get-IpLookupCountryName -CountryCode $ipInfoCountryCode
        }

        $locationParts = @(
            $ipInfoCity
            $ipInfoRegion
            $ipInfoCountry
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

        if ($locationParts.Count -gt 0) {
            $location = $locationParts -join ", "
        }

        $network = [string]$ipInfo.org

        if ([string]::IsNullOrWhiteSpace($network) -and -not [string]::IsNullOrWhiteSpace($ipInfoAsName)) {
            $network = "$ipInfoAsn $ipInfoAsName".Trim()

            if (-not [string]::IsNullOrWhiteSpace($ipInfoAsDomain)) {
                $network += " ($ipInfoAsDomain)"
            }
        }

        if ([string]::IsNullOrWhiteSpace($ipInfoAsn) -and $network -match '^(AS\d+)') {
            $ipInfoAsn = $matches[1]
        }

        if (
            [string]::IsNullOrWhiteSpace($ipInfoAsName) -and
            $network -match '^AS\d+\s+(.+)$'
        ) {
            $ipInfoAsName = $matches[1]

            if ($ipInfoAsName -match '^(.+?)\s+\(.+\)$') {
                $ipInfoAsName = $matches[1]
            }
        }

        $hostname = [string]$ipInfo.hostname
        $providerResponses++

        $ipInfoCountryDisplay = $ipInfoCountry
        if (-not [string]::IsNullOrWhiteSpace($ipInfoCountryCode)) {
            if ([string]::IsNullOrWhiteSpace($ipInfoCountryDisplay)) {
                $ipInfoCountryDisplay = $ipInfoCountryCode
            }
            elseif ($ipInfoCountryDisplay -notmatch "\($([regex]::Escape($ipInfoCountryCode))\)$") {
                $ipInfoCountryDisplay += " ($ipInfoCountryCode)"
            }
        }

        $ipInfoDetails = @("Status: $ipInfoDataTier data received.")

        $ipInfoLocationDisplay = @(
            $ipInfoCity
            $ipInfoRegion
            $ipInfoCountryDisplay
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

        if ($ipInfoLocationDisplay.Count -gt 0) {
            $ipInfoDetails += "Location: $($ipInfoLocationDisplay -join ', ')"
        }

        if (
            $null -ne $ipInfoLatitude -and
            $null -ne $ipInfoLongitude
        ) {
            $ipInfoDetails += "Coordinates: $ipInfoLatitude, $ipInfoLongitude"
        }

        $ipInfoContinentDisplay = $ipInfoContinent
        if (-not [string]::IsNullOrWhiteSpace($ipInfoContinentCode)) {
            if ([string]::IsNullOrWhiteSpace($ipInfoContinentDisplay)) {
                $ipInfoContinentDisplay = $ipInfoContinentCode
            }
            elseif ($ipInfoContinentDisplay -notmatch "\($([regex]::Escape($ipInfoContinentCode))\)$") {
                $ipInfoContinentDisplay += " ($ipInfoContinentCode)"
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($ipInfoContinentDisplay)) {
            $ipInfoDetails += "Continent: $ipInfoContinentDisplay"
        }

        if (-not [string]::IsNullOrWhiteSpace($ipInfoPostalCode)) {
            $ipInfoDetails += "Postal code: $ipInfoPostalCode"
        }

        if (-not [string]::IsNullOrWhiteSpace($ipInfoTimeZone)) {
            $ipInfoDetails += "Time zone: $ipInfoTimeZone"
        }

        if (-not [string]::IsNullOrWhiteSpace($hostname)) {
            $ipInfoDetails += "Domain name / reverse DNS: $hostname"
        }

        if (-not [string]::IsNullOrWhiteSpace($ipInfoAsn)) {
            $ipInfoDetails += "ASN: $ipInfoAsn"
        }

        if (-not [string]::IsNullOrWhiteSpace($ipInfoAsName)) {
            $ipInfoDetails += "Organization: $ipInfoAsName"
        }

        if (-not [string]::IsNullOrWhiteSpace($ipInfoAsDomain)) {
            $ipInfoDetails += "Organization domain: $ipInfoAsDomain"
        }

        if (-not [string]::IsNullOrWhiteSpace($ipInfoNetworkRange)) {
            $ipInfoDetails += "Network range: $ipInfoNetworkRange"
        }

        if (-not [string]::IsNullOrWhiteSpace($ipInfoAsType)) {
            $ipInfoDetails += "ASN type: $ipInfoAsType"
        }

        $ipInfoFlagFields = [ordered]@{
            "Anycast"                 = "anycast"
            "Anonymous infrastructure" = "is_anonymous"
            "Hosting/data center"     = "is_hosting"
            "Mobile carrier"          = "is_mobile"
            "Satellite provider"      = "is_satellite"
            "Bogon/reserved"          = "bogon"
        }

        $ipInfoThreatFlags = @()

        foreach ($flagLabel in $ipInfoFlagFields.Keys) {
            $flagPropertyName = $ipInfoFlagFields[$flagLabel]
            $flagProperty = $ipInfo.PSObject.Properties[$flagPropertyName]

            if ($null -ne $flagProperty -and $null -ne $flagProperty.Value) {
                $ipInfoThreatFlags += "$flagLabel=$(
                    ConvertTo-IpLookupYesNo -Value $flagProperty.Value
                )"

                if ($flagProperty.Value -eq $true) {
                    switch ($flagPropertyName) {
                    "anycast" {
                        $socSignals += "IPinfo marks the address as anycast; physical location may vary by requester."
                    }
                    "is_anonymous" {
                        $socSignals += "IPinfo flags anonymous infrastructure such as a VPN, proxy, Tor node, or relay."
                    }
                    "is_hosting" {
                        $socSignals += "IPinfo identifies hosting or data-center infrastructure."
                    }
                    "is_mobile" {
                        $socSignals += "IPinfo identifies a mobile-carrier address; attribution may be shared or temporary."
                    }
                    }
                }
            }
        }

        if (
            $null -eq $ipInfo.PSObject.Properties["anycast"] -and
            (Test-IpLookupValueReported -InputObject $ipInfo -PropertyName "is_anycast")
        ) {
            $ipInfoThreatFlags += "Anycast=$(ConvertTo-IpLookupYesNo -Value $ipInfo.is_anycast)"
        }

        if ($null -ne $ipInfo.anonymous) {
            $ipInfoAnonymousFields = [ordered]@{
                "VPN"   = "is_vpn"
                "Proxy" = "is_proxy"
                "Tor"   = "is_tor"
                "Relay" = "is_relay"
            }

            foreach ($flagLabel in $ipInfoAnonymousFields.Keys) {
                $flagPropertyName = $ipInfoAnonymousFields[$flagLabel]
                if (Test-IpLookupValueReported -InputObject $ipInfo.anonymous -PropertyName $flagPropertyName) {
                    $ipInfoThreatFlags += "$flagLabel=$(
                        ConvertTo-IpLookupYesNo -Value $ipInfo.anonymous.$flagPropertyName
                    )"
                }
            }
        }

        if ($null -ne $ipInfo.privacy) {
            $ipInfoPrivacyFields = [ordered]@{
                "VPN"     = "vpn"
                "Proxy"   = "proxy"
                "Tor"     = "tor"
                "Relay"   = "relay"
                "Hosting" = "hosting"
            }

            foreach ($flagLabel in $ipInfoPrivacyFields.Keys) {
                $flagPropertyName = $ipInfoPrivacyFields[$flagLabel]
                if (Test-IpLookupValueReported -InputObject $ipInfo.privacy -PropertyName $flagPropertyName) {
                    $flagText = "$flagLabel=$(ConvertTo-IpLookupYesNo -Value $ipInfo.privacy.$flagPropertyName)"
                    if ($ipInfoThreatFlags -notcontains $flagText) {
                        $ipInfoThreatFlags += $flagText
                    }
                }
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$ipInfo.privacy.service)) {
                $ipInfoThreatFlags += "Privacy service=$($ipInfo.privacy.service)"
            }
        }

        if ($ipInfoThreatFlags.Count -gt 0) {
            $ipInfoDetails += "Threat/context flags: $($ipInfoThreatFlags -join '; ')"
        }
    }
    catch {
        $ipInfoDetails = @("Status: Location and attribution request unavailable.")
    }

    # LevelBlue OTX supplies public threat-intelligence pulse associations.
    try {
        $otxHeaders = @{}
        $otxApiKey = Get-IpLookupApiKey -Name "OTX_API_KEY"

        if (-not [string]::IsNullOrWhiteSpace($otxApiKey)) {
            $otxHeaders["X-OTX-API-KEY"] = $otxApiKey
        }

        $otxIndicatorType = "IPv4"
        if ($ParsedAddress.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
            $otxIndicatorType = "IPv6"
        }

        $otxUri = "https://otx.alienvault.com/api/v1/indicators/$otxIndicatorType/$encodedIp/general"
        $otx = Invoke-IpLookupJsonRequest -Uri $otxUri -Headers $otxHeaders
        $otxPassiveDnsNames = @()

        try {
            $otxPassiveDnsUri = "https://otx.alienvault.com/api/v1/indicators/$otxIndicatorType/$encodedIp/passive_dns"
            $otxPassiveDns = Invoke-IpLookupJsonRequest -Uri $otxPassiveDnsUri -Headers $otxHeaders
            $otxPassiveDnsNames = @(
                @($otxPassiveDns.passive_dns) |
                    ForEach-Object {
                        ConvertTo-IpLookupSingleLine -Value $_.hostname -MaximumLength 120
                    } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique |
                    Select-Object -First 10
            )
        }
        catch {
            # Passive DNS is supplementary. Keep the general OTX result when
            # this endpoint has no history, is rate-limited, or is unavailable.
            $otxPassiveDnsNames = @()
        }

        $otxPulseCount = $null
        $otxHasPulseCount = $false
        $otxHasAssessment = $false

        if (Test-IpLookupValueReported -InputObject $otx.pulse_info -PropertyName "count") {
            $otxPulseCount = [int]$otx.pulse_info.count
            $otxHasPulseCount = $true
            $otxHasAssessment = $true
        }

        $threatSourcesChecked++
        $providerResponses++

        if ($otxHasPulseCount -and $otxPulseCount -gt 0) {
            $hasPossibleThreatSignal = $true
            $socSignals += "OTX links this IP to $otxPulseCount threat-intelligence pulse(s)."
        }

        if ([string]::IsNullOrWhiteSpace($location) -and -not [string]::IsNullOrWhiteSpace([string]$otx.country_name)) {
            $location = [string]$otx.country_name
        }

        if ([string]::IsNullOrWhiteSpace($network) -and -not [string]::IsNullOrWhiteSpace([string]$otx.asn)) {
            $network = [string]$otx.asn
        }

        $otxPulses = @($otx.pulse_info.pulses)
        $otxPulseNames = @(
            $otxPulses |
                ForEach-Object { ConvertTo-IpLookupSingleLine -Value $_.name -MaximumLength 120 } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique |
                Select-Object -First 5
        )
        $otxPulseTags = @(
            $otxPulses |
                ForEach-Object { $_.tags } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique |
                Select-Object -First 12
        )
        $otxAdversaries = @(
            $otxPulses |
                ForEach-Object { ConvertTo-IpLookupSingleLine -Value $_.adversary -MaximumLength 80 } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique |
                Select-Object -First 8
        )
        $otxIndustries = @(
            $otxPulses |
                ForEach-Object { $_.industries } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique |
                Select-Object -First 8
        )
        $otxTargetedCountries = @(
            $otxPulses |
                ForEach-Object { $_.targeted_countries } |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique |
                Select-Object -First 8
        )
        $otxReputation = ConvertTo-IpLookupSingleLine -Value $otx.reputation -MaximumLength 80

        if (-not [string]::IsNullOrWhiteSpace($otxReputation)) {
            $otxHasAssessment = $true
        }

        if ($otxHasAssessment) {
            $threatSourcesWithAssessment++
        }

        $otxDetails = @("Status: Threat-intelligence data received.")
        $otxCountryName = [string]$otx.country_name
        $otxCountryCode = [string]$otx.country_code
        $otxLocationParts = @(
            [string]$otx.city
            [string]$otx.region
            $otxCountryName
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

        if ($otxLocationParts.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($otxCountryCode)) {
            $otxLocationParts = @(Get-IpLookupCountryName -CountryCode $otxCountryCode)
        }

        if ($otxLocationParts.Count -gt 0) {
            $otxLocationDisplay = $otxLocationParts -join ", "
            if (
                -not [string]::IsNullOrWhiteSpace($otxCountryCode) -and
                $otxLocationDisplay -notmatch "\($([regex]::Escape($otxCountryCode))\)$"
            ) {
                $otxLocationDisplay += " ($otxCountryCode)"
            }
            $otxDetails += "Location: $otxLocationDisplay"
        }

        if (
            (Test-IpLookupValueReported -InputObject $otx -PropertyName "latitude") -and
            (Test-IpLookupValueReported -InputObject $otx -PropertyName "longitude")
        ) {
            $otxDetails += "Coordinates: $($otx.latitude), $($otx.longitude)"
        }

        if ($otxPassiveDnsNames.Count -gt 0) {
            $otxDetails += "Domain names / passive DNS (up to 10): $($otxPassiveDnsNames -join ', ')"
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$otx.asn)) {
            $otxDetails += "ASN / network owner: $($otx.asn)"
        }

        $otxThreatParts = @()
        if ($otxHasPulseCount) {
            $otxThreatParts += "Related pulses=$otxPulseCount"
        }

        if (-not [string]::IsNullOrWhiteSpace($otxReputation)) {
            $otxThreatParts += "Reputation=$otxReputation"
        }

        if ($otxThreatParts.Count -gt 0) {
            $otxDetails += "Threat detection: $($otxThreatParts -join '; ')"
        }

        if ($otxPulseNames.Count -gt 0) {
            $otxDetails += "Pulse names (up to 5): $($otxPulseNames -join ' | ')"
        }

        if ($otxPulseTags.Count -gt 0) {
            $otxDetails += "Tags (up to 12): $($otxPulseTags -join ', ')"
        }

        if ($otxAdversaries.Count -gt 0) {
            $otxDetails += "Adversaries: $($otxAdversaries -join ', ')"
        }

        if ($otxIndustries.Count -gt 0) {
            $otxDetails += "Industries: $($otxIndustries -join ', ')"
        }

        if ($otxTargetedCountries.Count -gt 0) {
            $otxDetails += "Targeted countries: $($otxTargetedCountries -join ', ')"
        }
    }
    catch {
        $otxDetails = @("Status: Threat-intelligence request unavailable.")
    }

    # VirusTotal is queried only when the user has configured an API key.
    $virusTotalApiKey = Get-IpLookupApiKey -Name "VIRUSTOTAL_API_KEY"
    if (-not [string]::IsNullOrWhiteSpace($virusTotalApiKey)) {
        try {
            $virusTotalHeaders = @{ "x-apikey" = $virusTotalApiKey }
            $virusTotalUri = "https://www.virustotal.com/api/v3/ip_addresses/$encodedIp"
            $virusTotal = Invoke-IpLookupJsonRequest -Uri $virusTotalUri -Headers $virusTotalHeaders
            $virusTotalAttributes = $virusTotal.data.attributes
            $virusTotalStats = $virusTotalAttributes.last_analysis_stats
            $virusTotalHasMalicious = Test-IpLookupValueReported -InputObject $virusTotalStats -PropertyName "malicious"
            $virusTotalHasSuspicious = Test-IpLookupValueReported -InputObject $virusTotalStats -PropertyName "suspicious"
            $virusTotalHasReputation = Test-IpLookupValueReported -InputObject $virusTotalAttributes -PropertyName "reputation"
            $virusTotalMalicious = $null
            $virusTotalSuspicious = $null
            $virusTotalReputation = $null

            if ($virusTotalHasMalicious) { $virusTotalMalicious = [int]$virusTotalStats.malicious }
            if ($virusTotalHasSuspicious) { $virusTotalSuspicious = [int]$virusTotalStats.suspicious }
            if ($virusTotalHasReputation) { $virusTotalReputation = [int]$virusTotalAttributes.reputation }

            $threatSourcesChecked++
            $providerResponses++

            if ($virusTotalHasMalicious -or $virusTotalHasSuspicious -or $virusTotalHasReputation) {
                $threatSourcesWithAssessment++
            }

            if ($virusTotalHasMalicious -and $virusTotalMalicious -ge 5) {
                $hasHighThreatSignal = $true
            }
            elseif (
                ($virusTotalHasMalicious -and $virusTotalMalicious -gt 0) -or
                ($virusTotalHasSuspicious -and $virusTotalSuspicious -gt 0)
            ) {
                $hasPossibleThreatSignal = $true
            }

            if ($virusTotalHasReputation -and $virusTotalReputation -lt 0) {
                $hasPossibleThreatSignal = $true
                $socSignals += "VirusTotal community reputation is negative ($virusTotalReputation)."
            }

            $virusTotalPositiveVerdicts = @()

            if ($virusTotalHasMalicious -and $virusTotalMalicious -gt 0) {
                $virusTotalPositiveVerdicts += "$virusTotalMalicious malicious"
            }

            if ($virusTotalHasSuspicious -and $virusTotalSuspicious -gt 0) {
                $virusTotalPositiveVerdicts += "$virusTotalSuspicious suspicious"
            }

            if ($virusTotalPositiveVerdicts.Count -gt 0) {
                $socSignals += "VirusTotal reports $($virusTotalPositiveVerdicts -join ' and ') engine result(s)."
            }

            $virusTotalFlaggingEngines = @()

            if (Test-IpLookupValueReported -InputObject $virusTotalAttributes -PropertyName "last_analysis_results") {
                $virusTotalFlaggingEngines = @(
                    foreach ($engineProperty in $virusTotalAttributes.last_analysis_results.PSObject.Properties) {
                        $engineCategory = [string]$engineProperty.Value.category

                        if ($engineCategory -eq "malicious" -or $engineCategory -eq "suspicious") {
                            ConvertTo-IpLookupSingleLine `
                                -Value $engineProperty.Value.engine_name `
                                -MaximumLength 60
                        }
                    }
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique |
                    Select-Object -First 12
            }

            $virusTotalTags = @()

            if (Test-IpLookupValueReported -InputObject $virusTotalAttributes -PropertyName "tags") {
                $virusTotalTags = @(
                    $virusTotalAttributes.tags |
                    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                    Sort-Object -Unique |
                    Select-Object -First 12
                )
            }

            $virusTotalLastAnalysis = $null

            if (Test-IpLookupValueReported -InputObject $virusTotalAttributes -PropertyName "last_analysis_date") {
                $virusTotalLastAnalysis = ConvertFrom-IpLookupUnixTime `
                    -Value $virusTotalAttributes.last_analysis_date
            }

            $virusTotalDetails = @("Status: Reputation data received.")
            $virusTotalCountryCode = [string]$virusTotalAttributes.country
            $virusTotalCountryName = Get-IpLookupCountryName -CountryCode $virusTotalCountryCode
            $virusTotalContinent = [string]$virusTotalAttributes.continent
            $virusTotalLocationParts = @(
                $virusTotalCountryName
                $virusTotalContinent
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }

            if ($virusTotalLocationParts.Count -gt 0) {
                $virusTotalLocationDisplay = $virusTotalLocationParts -join ", "
                if (
                    -not [string]::IsNullOrWhiteSpace($virusTotalCountryCode) -and
                    $virusTotalLocationDisplay -notmatch "\($([regex]::Escape($virusTotalCountryCode))\)"
                ) {
                    $virusTotalLocationDisplay += " ($virusTotalCountryCode)"
                }
                $virusTotalDetails += "Location: $virusTotalLocationDisplay"
            }

            $virusTotalDomainNames = @()
            if (Test-IpLookupValueReported -InputObject $virusTotalAttributes -PropertyName "last_dns_records") {
                $virusTotalDomainNames = @(
                    foreach ($dnsRecord in @($virusTotalAttributes.last_dns_records)) {
                        $dnsValue = ConvertTo-IpLookupSingleLine -Value $dnsRecord.value -MaximumLength 150
                        $dnsParsedAddress = $null

                        if (
                            -not [string]::IsNullOrWhiteSpace($dnsValue) -and
                            -not [System.Net.IPAddress]::TryParse($dnsValue, [ref]$dnsParsedAddress)
                        ) {
                            $dnsValue
                        }
                    }
                ) | Sort-Object -Unique | Select-Object -First 10
            }

            $virusTotalCertificateName = $null
            if (
                $null -ne $virusTotalAttributes.last_https_certificate -and
                $null -ne $virusTotalAttributes.last_https_certificate.subject
            ) {
                $virusTotalCertificateName = ConvertTo-IpLookupSingleLine `
                    -Value $virusTotalAttributes.last_https_certificate.subject.CN `
                    -MaximumLength 150
            }

            if ($virusTotalDomainNames.Count -gt 0) {
                $virusTotalDetails += "Domain names / recent DNS (up to 10): $($virusTotalDomainNames -join ', ')"
            }

            if (-not [string]::IsNullOrWhiteSpace($virusTotalCertificateName)) {
                $virusTotalDetails += "HTTPS certificate name: $virusTotalCertificateName"
            }

            if (Test-IpLookupValueReported -InputObject $virusTotalAttributes -PropertyName "asn") {
                $virusTotalDetails += "ASN: AS$($virusTotalAttributes.asn)"
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$virusTotalAttributes.as_owner)) {
                $virusTotalDetails += "Network owner: $($virusTotalAttributes.as_owner)"
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$virusTotalAttributes.network)) {
                $virusTotalDetails += "Network range: $($virusTotalAttributes.network)"
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$virusTotalAttributes.regional_internet_registry)) {
                $virusTotalDetails += "Regional registry: $($virusTotalAttributes.regional_internet_registry)"
            }

            $virusTotalVerdictParts = @()
            $virusTotalVerdictFields = [ordered]@{
                "Malicious"  = "malicious"
                "Suspicious" = "suspicious"
                "Harmless"   = "harmless"
                "Undetected" = "undetected"
                "Timeout"    = "timeout"
            }

            foreach ($verdictLabel in $virusTotalVerdictFields.Keys) {
                $verdictPropertyName = $virusTotalVerdictFields[$verdictLabel]
                if (Test-IpLookupValueReported -InputObject $virusTotalStats -PropertyName $verdictPropertyName) {
                    $virusTotalVerdictParts += "$verdictLabel=$($virusTotalStats.$verdictPropertyName)"
                }
            }

            if ($virusTotalVerdictParts.Count -gt 0) {
                $virusTotalDetails += "Threat detection: $($virusTotalVerdictParts -join '; ')"
            }

            if ($virusTotalHasReputation) {
                $virusTotalDetails += "Community reputation: $virusTotalReputation (negative is suspicious; positive is favorable)"
            }

            if ($null -ne $virusTotalAttributes.total_votes) {
                $virusTotalVoteParts = @()
                if (Test-IpLookupValueReported -InputObject $virusTotalAttributes.total_votes -PropertyName "malicious") {
                    $virusTotalVoteParts += "Malicious=$($virusTotalAttributes.total_votes.malicious)"
                }
                if (Test-IpLookupValueReported -InputObject $virusTotalAttributes.total_votes -PropertyName "harmless") {
                    $virusTotalVoteParts += "Harmless=$($virusTotalAttributes.total_votes.harmless)"
                }
                if ($virusTotalVoteParts.Count -gt 0) {
                    $virusTotalDetails += "Community votes: $($virusTotalVoteParts -join '; ')"
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($virusTotalLastAnalysis)) {
                $virusTotalDetails += "Last analysis: $virusTotalLastAnalysis"
            }

            if ($virusTotalFlaggingEngines.Count -gt 0) {
                $virusTotalDetails += "Flagging engines (up to 12): $($virusTotalFlaggingEngines -join ', ')"
            }

            if ($virusTotalTags.Count -gt 0) {
                $virusTotalDetails += "Tags (up to 12): $($virusTotalTags -join ', ')"
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$virusTotalAttributes.jarm)) {
                $virusTotalDetails += "JARM fingerprint: $($virusTotalAttributes.jarm)"
            }

            if ([string]::IsNullOrWhiteSpace($network) -and -not [string]::IsNullOrWhiteSpace([string]$virusTotalAttributes.as_owner)) {
                $network = [string]$virusTotalAttributes.as_owner

                if (
                    (Test-IpLookupValueReported -InputObject $virusTotalAttributes -PropertyName "asn") -and
                    [int]$virusTotalAttributes.asn -gt 0
                ) {
                    $network = "AS$($virusTotalAttributes.asn) $network"
                }
            }

            if ([string]::IsNullOrWhiteSpace($location) -and -not [string]::IsNullOrWhiteSpace([string]$virusTotalAttributes.country)) {
                $location = Get-IpLookupCountryName -CountryCode ([string]$virusTotalAttributes.country)
            }
        }
        catch {
            $virusTotalDetails = @(
                "Status: Request unavailable; check VIRUSTOTAL_API_KEY or its quota."
            )
        }
    }

    # AbuseIPDB is queried only when the user has configured an API key.
    $abuseIpDbApiKey = Get-IpLookupApiKey -Name "ABUSEIPDB_API_KEY"
    if (-not [string]::IsNullOrWhiteSpace($abuseIpDbApiKey)) {
        try {
            $abuseIpDbHeaders = @{
                "Key"    = $abuseIpDbApiKey
                "Accept" = "application/json"
            }
            # The standard CHECK response already includes the SOC fields used
            # below. Omitting verbose avoids downloading thousands of raw reports.
            $abuseIpDbUri = "https://api.abuseipdb.com/api/v2/check?ipAddress=$encodedIp&maxAgeInDays=90"
            $abuseIpDb = Invoke-IpLookupJsonRequest -Uri $abuseIpDbUri -Headers $abuseIpDbHeaders
            $abuseData = $abuseIpDb.data
            $abuseHasScore = Test-IpLookupValueReported -InputObject $abuseData -PropertyName "abuseConfidenceScore"
            $abuseHasReports = Test-IpLookupValueReported -InputObject $abuseData -PropertyName "totalReports"
            $abuseHasDistinctUsers = Test-IpLookupValueReported -InputObject $abuseData -PropertyName "numDistinctUsers"
            $abuseScore = $null
            $abuseReports = $null
            $abuseDistinctUsers = $null

            if ($abuseHasScore) { $abuseScore = [int]$abuseData.abuseConfidenceScore }
            if ($abuseHasReports) { $abuseReports = [int]$abuseData.totalReports }
            if ($abuseHasDistinctUsers) { $abuseDistinctUsers = [int]$abuseData.numDistinctUsers }

            $threatSourcesChecked++
            $providerResponses++

            if ($abuseHasScore -or $abuseHasReports) {
                $threatSourcesWithAssessment++
            }

            if ($abuseHasScore -and $abuseScore -ge 75) {
                $hasHighThreatSignal = $true
            }
            elseif (
                ($abuseHasScore -and $abuseScore -ge 25) -or
                ($abuseHasReports -and $abuseReports -gt 0)
            ) {
                $hasPossibleThreatSignal = $true
            }

            if (
                ($abuseHasReports -and $abuseReports -gt 0) -or
                ($abuseHasScore -and $abuseScore -gt 0)
            ) {
                $abuseSignalParts = @()

                if ($abuseHasScore) { $abuseSignalParts += "$abuseScore% confidence" }
                if ($abuseHasReports) { $abuseSignalParts += "$abuseReports report(s) in 90 days" }
                if ($abuseHasDistinctUsers) { $abuseSignalParts += "$abuseDistinctUsers distinct reporter(s)" }

                $socSignals += "AbuseIPDB reports $($abuseSignalParts -join ', ')."
            }

            if (
                (Test-IpLookupValueReported -InputObject $abuseData -PropertyName "isTor") -and
                $abuseData.isTor -eq $true
            ) {
                $socSignals += "AbuseIPDB identifies this address as a Tor exit node."
            }

            $abuseCountryCode = [string]$abuseData.countryCode
            $abuseCountryName = Get-IpLookupCountryName -CountryCode $abuseCountryCode
            $abuseHostnames = @(
                $abuseData.hostnames |
                    ForEach-Object { ConvertTo-IpLookupSingleLine -Value $_ -MaximumLength 100 } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique |
                    Select-Object -First 3
            )
            $abuseLastReported = $null

            if (Test-IpLookupValueReported -InputObject $abuseData -PropertyName "lastReportedAt") {
                $abuseLastReported = ConvertTo-IpLookupUtcTime -Value $abuseData.lastReportedAt
            }

            $abuseIpDbDetails = @("Status: Abuse-report data received.")
            if (-not [string]::IsNullOrWhiteSpace($abuseCountryName)) {
                $abuseLocationDisplay = $abuseCountryName
                if (
                    -not [string]::IsNullOrWhiteSpace($abuseCountryCode) -and
                    $abuseLocationDisplay -notmatch "\($([regex]::Escape($abuseCountryCode))\)$"
                ) {
                    $abuseLocationDisplay += " ($abuseCountryCode)"
                }
                $abuseIpDbDetails += "Location: $abuseLocationDisplay"
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$abuseData.domain)) {
                $abuseIpDbDetails += "Domain name: $($abuseData.domain)"
            }

            if ($abuseHostnames.Count -gt 0) {
                $abuseIpDbDetails += "Hostnames (up to 3): $($abuseHostnames -join ', ')"
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$abuseData.isp)) {
                $abuseIpDbDetails += "Network owner / ISP: $($abuseData.isp)"
            }

            if (-not [string]::IsNullOrWhiteSpace([string]$abuseData.usageType)) {
                $abuseIpDbDetails += "Usage type: $($abuseData.usageType)"
            }

            $abuseThreatParts = @()

            if ($abuseHasScore) {
                $abuseThreatParts += "Abuse confidence=$abuseScore%"
            }

            if ($abuseHasReports) {
                $abuseThreatParts += "Reports in 90 days=$abuseReports"
            }

            if ($abuseHasDistinctUsers) {
                $abuseThreatParts += "Distinct reporters=$abuseDistinctUsers"
            }

            if (
                Test-IpLookupValueReported -InputObject $abuseData -PropertyName "isTor"
            ) {
                $abuseThreatParts += "Tor exit node=$(ConvertTo-IpLookupYesNo -Value $abuseData.isTor)"
            }

            if (
                Test-IpLookupValueReported -InputObject $abuseData -PropertyName "isWhitelisted"
            ) {
                $abuseThreatParts += "Whitelisted=$(ConvertTo-IpLookupYesNo -Value $abuseData.isWhitelisted)"
            }

            if ($abuseThreatParts.Count -gt 0) {
                $abuseIpDbDetails += "Threat detection: $($abuseThreatParts -join '; ')"
            }

            if (-not [string]::IsNullOrWhiteSpace($abuseLastReported)) {
                $abuseIpDbDetails += "Last reported: $abuseLastReported"
            }

            $abuseAddressFacts = @()
            if (Test-IpLookupValueReported -InputObject $abuseData -PropertyName "ipVersion") {
                $abuseAddressFacts += "IPv$($abuseData.ipVersion)"
            }
            if (Test-IpLookupValueReported -InputObject $abuseData -PropertyName "isPublic") {
                $abuseAddressFacts += "Public=$(ConvertTo-IpLookupYesNo -Value $abuseData.isPublic)"
            }
            if ($abuseAddressFacts.Count -gt 0) {
                $abuseIpDbDetails += "Address classification: $($abuseAddressFacts -join '; ')"
            }

            if ([string]::IsNullOrWhiteSpace($network) -and -not [string]::IsNullOrWhiteSpace([string]$abuseData.isp)) {
                $network = [string]$abuseData.isp

                if (-not [string]::IsNullOrWhiteSpace([string]$abuseData.domain)) {
                    $network += " ($($abuseData.domain))"
                }
            }

            if ([string]::IsNullOrWhiteSpace($location) -and -not [string]::IsNullOrWhiteSpace([string]$abuseData.countryCode)) {
                $location = Get-IpLookupCountryName -CountryCode ([string]$abuseData.countryCode)
            }
        }
        catch {
            $abuseIpDbDetails = @(
                "Status: Request unavailable; check ABUSEIPDB_API_KEY or its quota."
            )
        }
    }

    if ($hasHighThreatSignal) {
        $overall = "HIGH-RISK SIGNALS"
    }
    elseif ($hasPossibleThreatSignal) {
        $overall = "POSSIBLE THREAT SIGNALS"
    }
    elseif ($threatSourcesWithAssessment -gt 0) {
        $overall = "NO THREAT FLAGS FOUND"
    }
    else {
        $overall = "LIMITED THREAT DATA"
    }

    if ($socSignals.Count -eq 0) {
        if ($threatSourcesWithAssessment -gt 0) {
            $socSignals = @("No positive threat signals were returned by the responding threat feeds.")
        }
        elseif ($threatSourcesChecked -gt 0) {
            $socSignals = @("Threat feeds responded but did not return a usable threat assessment.")
        }
        else {
            $socSignals = @("No threat feed responded; review the provider status lines below.")
        }
    }

    # Tab 4 uses the compact SOC view. Each provider remains separate, but
    # secondary enrichment fields are removed so the analyst can scan the
    # location, ownership, domain, and threat findings quickly.
    if ($ImportantOnly) {
        $ipInfoDetails = @(
            $ipInfoDetails | Where-Object {
                [string]$_ -match '^(Status|Location|Domain name / reverse DNS|ASN|Organization|Organization domain|Network range|Threat/context flags):'
            }
        )

        $otxDetails = @(
            $otxDetails | Where-Object {
                [string]$_ -match '^(Status|Location|Domain names / passive DNS|ASN / network owner|Threat detection|Pulse names|Tags):'
            }
        )

        $virusTotalDetails = @(
            $virusTotalDetails | Where-Object {
                [string]$_ -match '^(Status|Location|Domain names / recent DNS|HTTPS certificate name|ASN|Network owner|Network range|Threat detection|Community reputation|Last analysis|Flagging engines|Tags):'
            }
        )

        $abuseIpDbDetails = @(
            $abuseIpDbDetails | Where-Object {
                [string]$_ -match '^(Status|Location|Domain name|Hostnames|Network owner / ISP|Usage type|Threat detection|Last reported):'
            }
        )
    }

    $summaryLines = @(
        "IP: $IpAddress"
        "OVERALL: $overall"
        "Providers responding: $providerResponses of 4"
        "Threat feeds checked: $threatSourcesChecked of 3"
        ""
        "QUICK SOC FLAGS / CONTEXT"
        $socSignals
    )

    $combinedAttributionDetails = @()

    if (-not [string]::IsNullOrWhiteSpace($location)) {
        $combinedAttributionDetails += "Location: $location"
    }

    if (-not [string]::IsNullOrWhiteSpace($network)) {
        $combinedAttributionDetails += "Network/owner: $network"
    }

    if (-not [string]::IsNullOrWhiteSpace($hostname)) {
        $combinedAttributionDetails += "Hostname: $hostname"
    }

    if (-not $ImportantOnly -and $combinedAttributionDetails.Count -gt 0) {
        $summaryLines += ""
        $summaryLines += "COMBINED ATTRIBUTION"
        $summaryLines += $combinedAttributionDetails
    }

    $summaryLines += @(
        ""
        "1. IPINFO"
        $ipInfoDetails
        ""
        "2. LEVELBLUE OTX"
        $otxDetails
        ""
        "3. VIRUSTOTAL"
        $virusTotalDetails
        ""
        "4. ABUSEIPDB"
        $abuseIpDbDetails
        ""
        "Analyst note: Reputation is point-in-time context. Corroborate it with the alert, logs, asset role, direction of traffic, and observed behavior."
    )

    return $summaryLines -join $newLine
}

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

# Creates a second button that builds the same automatic summary without
# opening any browser tabs.
$ipSummaryOnlyButton = New-Object System.Windows.Forms.Button
$ipSummaryOnlyButton.Text = "Summary Only"
$ipSummaryOnlyButton.Size = New-Object System.Drawing.Size(100,30)
$ipSummaryOnlyButton.Location = New-Object System.Drawing.Point(285,82)


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

# ------------------------------------------------------------
# THREAT AND LOCATION SUMMARY PANEL
# ------------------------------------------------------------

$ipSummaryLabel = New-Object System.Windows.Forms.Label
$ipSummaryLabel.Text = "Threat & Location Summary:"
$ipSummaryLabel.UseMnemonic = $false
$ipSummaryLabel.AutoSize = $true
$ipSummaryLabel.Location = New-Object System.Drawing.Point(400,25)

$ipCopySummaryButton = New-Object System.Windows.Forms.Button
$ipCopySummaryButton.Text = "Copy Summary"
$ipCopySummaryButton.Size = New-Object System.Drawing.Size(100,28)
$ipCopySummaryButton.Location = New-Object System.Drawing.Point(590,18)
$ipCopySummaryButton.Enabled = $false

$ipSummaryBox = New-Object System.Windows.Forms.RichTextBox
$ipSummaryBox.Location = New-Object System.Drawing.Point(400,50)
$ipSummaryBox.Size = New-Object System.Drawing.Size(290,460)
$ipSummaryBox.ReadOnly = $true
$ipSummaryBox.WordWrap = $true
$ipSummaryBox.ScrollBars = "Vertical"
$ipSummaryBox.DetectUrls = $false
$ipSummaryBox.Text = "Run a lookup to see where the IP is from and whether threat-intelligence sources report it."
$ipSummaryBox.Anchor = "Top, Bottom, Left, Right"

$ipCopySummaryButton.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($ipSummaryBox.Text)) {
        [System.Windows.Forms.Clipboard]::SetText($ipSummaryBox.Text)
    }
})

# Chrome cleanup instructions shown directly below the website checklist
$ipCloseTabsTip = New-Object System.Windows.Forms.Label
$ipCloseTabsTip.Text = 'Chrome tip: To close all lookup tabs, right-click the tab' + "`n" + 'you want to keep, then choose "Close tabs to the right."'
$ipCloseTabsTip.AutoSize = $false
$ipCloseTabsTip.Size = New-Object System.Drawing.Size(455,36)
$ipCloseTabsTip.Location = New-Object System.Drawing.Point(20,514)
$ipCloseTabsTip.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)

# Explains which services contribute to the automatic summary
$ipSummarySourcesTip = New-Object System.Windows.Forms.Label
$ipSummarySourcesTip.Text = "The four websites that work with the summary are:`n1. IPinfo`n2. LevelBlue OTX`n3. VirusTotal`n4. AbuseIPDB"
$ipSummarySourcesTip.AutoSize = $false
$ipSummarySourcesTip.Size = New-Object System.Drawing.Size(455,82)
$ipSummarySourcesTip.Location = New-Object System.Drawing.Point(20,552)
$ipSummarySourcesTip.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)


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

# Summary-only button styling
$ipSummaryOnlyButton.FlatStyle = "Flat"
$ipSummaryOnlyButton.BackColor = [System.Drawing.Color]::FromArgb(40,145,120)
$ipSummaryOnlyButton.ForeColor = [System.Drawing.Color]::White
$ipSummaryOnlyButton.FlatAppearance.BorderSize = 0
$ipSummaryOnlyButton.FlatAppearance.MouseOverBackColor = `
    [System.Drawing.Color]::FromArgb(30,125,100)
$ipSummaryOnlyButton.FlatAppearance.MouseDownBackColor = `
    [System.Drawing.Color]::FromArgb(20,105,85)



# ------------------------------------------------------------
# PRIVATE IP REFERENCE
# ------------------------------------------------------------

# Gives the private IP section a warm orange color
$ipPrivateLabel.ForeColor = [System.Drawing.Color]::FromArgb(190,90,35)

# Makes the Chrome tab-cleanup instruction readable without overpowering the form
$ipCloseTabsTip.ForeColor = [System.Drawing.Color]::FromArgb(75,85,95)
$ipSummarySourcesTip.ForeColor = [System.Drawing.Color]::FromArgb(75,85,95)

# Threat/location results styling
$ipSummaryLabel.ForeColor = [System.Drawing.Color]::FromArgb(35,95,160)
$ipSummaryBox.BackColor = [System.Drawing.Color]::White
$ipSummaryBox.ForeColor = [System.Drawing.Color]::FromArgb(35,35,35)
$ipSummaryBox.BorderStyle = "FixedSingle"
$ipSummaryBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$ipCopySummaryButton.FlatStyle = "Flat"
$ipCopySummaryButton.BackColor = [System.Drawing.Color]::FromArgb(90,105,120)
$ipCopySummaryButton.ForeColor = [System.Drawing.Color]::White
$ipCopySummaryButton.FlatAppearance.BorderSize = 0

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

# Both buttons use the same lookup routine. Run Lookup opens the checked
# websites and builds the summary; Summary Only skips the browser tabs.
$ipRunLookupAction = {
    param(
        [bool]$OpenWebsites
    )

    # Normalize raw IPv4/IPv6 values, URLs, ports, and pasted punctuation.
    $ipOriginalInput = $ipTextBox.Text
    $ipIp = ConvertTo-NormalizedIpAddress -InputText $ipOriginalInput

    # Updates the search box so the extracted canonical address is visible.
    $ipTextBox.Text = [string]$ipIp

    # Creates a variable PowerShell will use to validate the IP
    $ipValidIP = $null


    # --------------------------------------------------------
    # VALIDATE IP ADDRESS
    # --------------------------------------------------------

    # Checks whether the entered text is a valid IPv4 or IPv6 address
    if (
        [string]::IsNullOrWhiteSpace($ipIp) -or
        -not [System.Net.IPAddress]::TryParse($ipIp, [ref]$ipValidIP)
    ) {

        # Shows an error popup if the IP is invalid
        [System.Windows.Forms.MessageBox]::Show(
            "'$($ipOriginalInput.Trim())' does not contain a valid IPv4 or IPv6 address.",
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

    if ($OpenWebsites) {
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
    }

    # --------------------------------------------------------
    # BUILD THREAT AND LOCATION SUMMARY
    # --------------------------------------------------------

    # During a full lookup, browser tabs open before slow API responses run.
    $ipButton.Enabled = $false
    $ipSummaryOnlyButton.Enabled = $false
    $ipCopySummaryButton.Enabled = $false
    $ipSummaryBox.Text = "Checking public location and threat-intelligence sources..."
    $mainForm.UseWaitCursor = $true
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $ipSummaryBox.Text = Get-IpThreatLocationSummary `
            -IpAddress $ipIp `
            -ParsedAddress $ipValidIP
        $ipSummaryBox.SelectionStart = 0
        $ipSummaryBox.ScrollToCaret()
        $ipCopySummaryButton.Enabled = $true
    }
    catch {
        $ipSummaryBox.Text = (
            "IP: $ipIp`r`nOVERALL: SUMMARY UNAVAILABLE`r`n`r`n" +
            "The lookup websites can still be reviewed in Chrome."
        )
    }
    finally {
        $mainForm.UseWaitCursor = $false
        $ipButton.Enabled = $true
        $ipSummaryOnlyButton.Enabled = $true
    }
    


    # --------------------------------------------------------
    # RESET SEARCH BOX AFTER LOOKUP
    # --------------------------------------------------------

    # Clears the IP that was just searched
    $ipTextBox.Clear()

    # Places the cursor back in the search box
    $ipTextBox.Focus()
}

# The normal button and Enter key perform the complete lookup.
$ipButton.Add_Click({
    & $ipRunLookupAction -OpenWebsites $true
})

# This button only refreshes the automatic summary panel.
$ipSummaryOnlyButton.Add_Click({
    & $ipRunLookupAction -OpenWebsites $false
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

# Adds the Summary Only button
$ipTab.Controls.Add($ipSummaryOnlyButton)

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

# Adds the threat/location summary panel
$ipTab.Controls.Add($ipSummaryLabel)
$ipTab.Controls.Add($ipSummaryBox)
$ipTab.Controls.Add($ipCopySummaryButton)

# Adds the Chrome tab cleanup instructions
$ipTab.Controls.Add($ipCloseTabsTip)

# Adds the automatic-summary source explanation
$ipTab.Controls.Add($ipSummarySourcesTip)


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

# ============================================================
# CASE IP LOOKUPS TAB
# ============================================================

$caseIpInstructions = New-Object System.Windows.Forms.Label
$caseIpInstructions.Text = "Press Fill Airtable on the Stellar tab to collect and look up the case IP addresses."
$caseIpInstructions.AutoSize = $false
$caseIpInstructions.Size = New-Object System.Drawing.Size(660, 24)
$caseIpInstructions.Location = New-Object System.Drawing.Point(20, 18)
$caseIpInstructions.ForeColor = [System.Drawing.Color]::FromArgb(35, 45, 55)
$caseIpTab.Controls.Add($caseIpInstructions)

$caseIpPrivateLabel = New-Object System.Windows.Forms.Label
$caseIpPrivateLabel.Text = "Private / Reserved (0)"
$caseIpPrivateLabel.AutoSize = $true
$caseIpPrivateLabel.Location = New-Object System.Drawing.Point(20, 52)
$caseIpPrivateLabel.ForeColor = [System.Drawing.Color]::FromArgb(190, 90, 35)
$caseIpTab.Controls.Add($caseIpPrivateLabel)

$caseIpPublicLabel = New-Object System.Windows.Forms.Label
$caseIpPublicLabel.Text = "Public (0)"
$caseIpPublicLabel.AutoSize = $true
$caseIpPublicLabel.Location = New-Object System.Drawing.Point(365, 52)
$caseIpPublicLabel.ForeColor = [System.Drawing.Color]::FromArgb(35, 95, 160)
$caseIpTab.Controls.Add($caseIpPublicLabel)

$caseIpPrivateList = New-Object System.Windows.Forms.ListBox
$caseIpPrivateList.Size = New-Object System.Drawing.Size(315, 135)
$caseIpPrivateList.Location = New-Object System.Drawing.Point(20, 75)
$caseIpPrivateList.HorizontalScrollbar = $true
$caseIpPrivateList.BackColor = [System.Drawing.Color]::White
$caseIpPrivateList.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$caseIpTab.Controls.Add($caseIpPrivateList)

$caseIpPublicList = New-Object System.Windows.Forms.ListBox
$caseIpPublicList.Size = New-Object System.Drawing.Size(315, 135)
$caseIpPublicList.Location = New-Object System.Drawing.Point(365, 75)
$caseIpPublicList.HorizontalScrollbar = $true
$caseIpPublicList.BackColor = [System.Drawing.Color]::White
$caseIpPublicList.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$caseIpTab.Controls.Add($caseIpPublicList)

$caseIpRunButton = New-Object System.Windows.Forms.Button
$caseIpRunButton.Text = "Look Up All Public IPs"
$caseIpRunButton.Size = New-Object System.Drawing.Size(180, 32)
$caseIpRunButton.Location = New-Object System.Drawing.Point(20, 225)
$caseIpRunButton.FlatStyle = "Flat"
$caseIpRunButton.FlatAppearance.BorderSize = 0
$caseIpRunButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$caseIpRunButton.ForeColor = [System.Drawing.Color]::White
$caseIpRunButton.Enabled = $false
$caseIpTab.Controls.Add($caseIpRunButton)

$caseIpCopyButton = New-Object System.Windows.Forms.Button
$caseIpCopyButton.Text = "Copy Results"
$caseIpCopyButton.Size = New-Object System.Drawing.Size(110, 32)
$caseIpCopyButton.Location = New-Object System.Drawing.Point(210, 225)
$caseIpCopyButton.FlatStyle = "Flat"
$caseIpCopyButton.FlatAppearance.BorderSize = 0
$caseIpCopyButton.BackColor = [System.Drawing.Color]::FromArgb(90, 105, 120)
$caseIpCopyButton.ForeColor = [System.Drawing.Color]::White
$caseIpCopyButton.Enabled = $false
$caseIpTab.Controls.Add($caseIpCopyButton)

$caseIpStatusLabel = New-Object System.Windows.Forms.Label
$caseIpStatusLabel.Text = "No Stellar case text has been pasted yet."
$caseIpStatusLabel.AutoSize = $false
$caseIpStatusLabel.Size = New-Object System.Drawing.Size(345, 32)
$caseIpStatusLabel.Location = New-Object System.Drawing.Point(335, 225)
$caseIpStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$caseIpStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 95)
$caseIpTab.Controls.Add($caseIpStatusLabel)

$caseIpResultsBox = New-Object System.Windows.Forms.TextBox
$caseIpResultsBox.Location = New-Object System.Drawing.Point(20, 270)
$caseIpResultsBox.Size = New-Object System.Drawing.Size(660, 350)
$caseIpResultsBox.Multiline = $true
$caseIpResultsBox.ScrollBars = "Vertical"
$caseIpResultsBox.BackColor = [System.Drawing.Color]::White
$caseIpResultsBox.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$caseIpResultsBox.BorderStyle = "FixedSingle"
$caseIpResultsBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$caseIpResultsBox.Text = "Public-IP results from IPinfo, LevelBlue OTX, VirusTotal, and AbuseIPDB will appear here."
$caseIpTab.Controls.Add($caseIpResultsBox)

$caseIpAddressByDisplay = @{}

# Clicking either list copies the exact IPv4 or IPv6 address that was clicked.
$caseIpPrivateList.Add_MouseClick({
    param($listControl, $mouseClickEvent)

    $clickedIndex = $listControl.IndexFromPoint($mouseClickEvent.Location)
    if ($clickedIndex -ge 0) {
        $clickedDisplay = [string]$listControl.Items[$clickedIndex]
        $clickedAddress = [string]$caseIpAddressByDisplay[$clickedDisplay]

        if ([string]::IsNullOrWhiteSpace($clickedAddress)) {
            $displayAddressPart = ($clickedDisplay -split '\s+\[', 2)[0]
            $clickedAddress = ConvertTo-NormalizedIpAddress -InputText $displayAddressPart
        }

        if (-not [string]::IsNullOrWhiteSpace($clickedAddress)) {
            [System.Windows.Forms.Clipboard]::SetText($clickedAddress)
            $caseIpStatusLabel.Text = "Copied IP: $clickedAddress"
        }
    }
})

$caseIpPublicList.Add_MouseClick({
    param($listControl, $mouseClickEvent)

    $clickedIndex = $listControl.IndexFromPoint($mouseClickEvent.Location)
    if ($clickedIndex -ge 0) {
        $clickedDisplay = [string]$listControl.Items[$clickedIndex]
        $clickedAddress = [string]$caseIpAddressByDisplay[$clickedDisplay]

        if ([string]::IsNullOrWhiteSpace($clickedAddress)) {
            $displayAddressPart = ($clickedDisplay -split '\s+\[', 2)[0]
            $clickedAddress = ConvertTo-NormalizedIpAddress -InputText $displayAddressPart
        }

        if (-not [string]::IsNullOrWhiteSpace($clickedAddress)) {
            [System.Windows.Forms.Clipboard]::SetText($clickedAddress)
            $caseIpStatusLabel.Text = "Copied IP: $clickedAddress"
        }
    }
})

# Rebuild the two lists only when Fill Airtable is pressed. Exact duplicate
# addresses are removed, and private/reserved values are never sent to an API.
$caseIpRefreshAction = {
    param([string]$Text)

    $caseIpPrivateList.BeginUpdate()
    $caseIpPublicList.BeginUpdate()

    try {
        $caseIpPrivateList.Items.Clear()
        $caseIpPublicList.Items.Clear()
        $caseIpAddressByDisplay.Clear()

        foreach ($caseIpRecord in @(Get-IpAddressesFromText -Text $Text)) {
            $caseIpDisplay = [string]$caseIpRecord.Display
            $caseIpAddressByDisplay[$caseIpDisplay] = [string]$caseIpRecord.Address

            if ($caseIpRecord.IsPublic) {
                [void]$caseIpPublicList.Items.Add($caseIpDisplay)
            }
            else {
                [void]$caseIpPrivateList.Items.Add($caseIpDisplay)
            }
        }
    }
    finally {
        $caseIpPrivateList.EndUpdate()
        $caseIpPublicList.EndUpdate()
    }

    $caseIpPrivateLabel.Text = "Private / Reserved ($($caseIpPrivateList.Items.Count))"
    $caseIpPublicLabel.Text = "Public ($($caseIpPublicList.Items.Count))"
    $caseIpRunButton.Enabled = ($caseIpPublicList.Items.Count -gt 0)
    $caseIpCopyButton.Enabled = $false
    $caseIpResultsBox.Text = "Public-IP results from IPinfo, LevelBlue OTX, VirusTotal, and AbuseIPDB will appear here."

    if ([string]::IsNullOrWhiteSpace($Text)) {
        $caseIpStatusLabel.Text = "No Stellar case text has been pasted yet."
    }
    else {
        $caseIpTotalCount = $caseIpPrivateList.Items.Count + $caseIpPublicList.Items.Count
        $caseIpStatusLabel.Text = "$caseIpTotalCount unique IP address(es) found."
    }
}

$caseIpRunLookupAction = {
    if ($caseIpPublicList.Items.Count -eq 0) {
        return
    }

    # Always discard manual changes and rebuild the report from the APIs.
    $caseIpResultsBox.Clear()
    $caseIpRunButton.Enabled = $false
    $caseIpCopyButton.Enabled = $false
    $mainForm.UseWaitCursor = $true
    $caseIpResultSections = New-Object System.Collections.Generic.List[string]

    try {
        $caseIpLookupNumber = 0
        $caseIpLookupTotal = $caseIpPublicList.Items.Count

        foreach ($caseIpListItem in $caseIpPublicList.Items) {
            $caseIpLookupNumber++
            $caseIpDisplay = [string]$caseIpListItem
            $caseIpAddress = [string]$caseIpAddressByDisplay[$caseIpDisplay]

            if ([string]::IsNullOrWhiteSpace($caseIpAddress)) {
                $displayAddressPart = ($caseIpDisplay -split '\s+\[', 2)[0]
                $caseIpAddress = ConvertTo-NormalizedIpAddress -InputText $displayAddressPart
            }

            $caseIpStatusLabel.Text = "Looking up $caseIpLookupNumber of $caseIpLookupTotal`: $caseIpAddress"
            [System.Windows.Forms.Application]::DoEvents()

            $caseIpParsedAddress = $null
            if (-not [System.Net.IPAddress]::TryParse(
                $caseIpAddress,
                [ref]$caseIpParsedAddress
            )) {
                continue
            }

            try {
                $caseIpSummary = Get-IpThreatLocationSummary `
                    -IpAddress $caseIpAddress `
                    -ParsedAddress $caseIpParsedAddress `
                    -ImportantOnly
            }
            catch {
                $caseIpSummary = "IP: $caseIpAddress`r`nOVERALL: SUMMARY UNAVAILABLE"
            }

            [void]$caseIpResultSections.Add($caseIpSummary)
            $caseIpResultsBox.Text = $caseIpResultSections -join (
                "`r`n`r`n" + ("=" * 68) + "`r`n`r`n"
            )
            $caseIpResultsBox.SelectionStart = $caseIpResultsBox.Text.Length
            $caseIpResultsBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }

        $caseIpResultsBox.SelectionStart = 0
        $caseIpResultsBox.ScrollToCaret()
        $caseIpCopyButton.Enabled = ($caseIpResultSections.Count -gt 0)
        $caseIpStatusLabel.Text = "Finished $($caseIpResultSections.Count) public IP lookup(s)."
    }
    finally {
        $mainForm.UseWaitCursor = $false
        $caseIpRunButton.Enabled = ($caseIpPublicList.Items.Count -gt 0)
    }
}

$caseIpRunButton.Add_Click({
    & $caseIpRunLookupAction
})

$caseIpCopyButton.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($caseIpResultsBox.Text)) {
        [System.Windows.Forms.Clipboard]::SetText($caseIpResultsBox.Text)
    }
})

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

    # Collect the Ctrl+A case IPs at the moment the analyst submits the case,
    # then show the fourth tab while the rest of the Airtable workflow runs.
    & $caseIpRefreshAction -Text $stellarText
    $tabControl.SelectedTab = $caseIpTab
    [System.Windows.Forms.Application]::DoEvents()

    # -----------------------------
    # JSON DATA
    # -----------------------------

    $stellarJsonText = ""
    $stellarJsonSourceIP = ""
    $stellarJsonDestinationIP = ""
    $stellarJsonHostIP = ""
    $stellarJsonUser = ""
    $stellarAlertDescription = ""

    # Isolate one valid JSON value without capturing Overview or another object.
    # Keep the JSON's original indentation and line breaks for readability.
    $stellarJsonText = Get-JsonFromText -Text $stellarText

    # Get the same dynamic paragraph Stellar displays on the Overview tab.
    $stellarAlertDescription = Get-StellarAlertDescription -JsonText $stellarJsonText

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

    # User fallback for alerts whose Case Score Breakdown did not load.
    # Prefer engid_name, then use computer_name when engid_name is unavailable.
    if ($stellarJsonText -match '(?i)"engid_name"\s*:\s*"([^"]+)"') {
        $stellarJsonUser = $matches[1].Trim()
    }
    elseif ($stellarJsonText -match '(?i)"computer_name"\s*:\s*"([^"]+)"') {
        $stellarJsonUser = $matches[1].Trim()
    }

    # -----------------------------
    # DETECTION TYPE
    # -----------------------------

    $stellarDetectionType = ""

    
    if ($stellarText -match '(?i)Sophos') {
        $stellarDetectionType = "EDR Alert"
    }
    elseif ($stellarText -match 'XDR') {
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

    # If the visible page did not include Run Analysis, use Stellar's official
    # JSON display-name field as the Issue. This only runs while Issue is blank.
    if (-not $stellarIssue -and $stellarJsonText) {
        try {
            $stellarIssueJson = ConvertFrom-Json `
                -InputObject $stellarJsonText `
                -ErrorAction Stop

            foreach ($stellarIssueJsonRecord in @($stellarIssueJson)) {
                # Support both a flattened xdr_event.display_name property and
                # an xdr_event object containing a display_name property.
                $stellarJsonDisplayName =
                    $stellarIssueJsonRecord.'xdr_event.display_name'

                if (
                    [string]::IsNullOrWhiteSpace([string]$stellarJsonDisplayName) -and
                    $null -ne $stellarIssueJsonRecord.xdr_event
                ) {
                    $stellarJsonDisplayName =
                        $stellarIssueJsonRecord.xdr_event.display_name
                }

                if (-not [string]::IsNullOrWhiteSpace([string]$stellarJsonDisplayName)) {
                    $stellarIssue = (
                        [string]$stellarJsonDisplayName -replace '\s+', ' '
                    ).Trim()
                    break
                }
            }
        }
        catch {
            # Keep trying the visible Issue-label fallback below.
        }
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

    if ([string]::IsNullOrWhiteSpace($stellarCaseScoreBreakdown)) {
        # Critical alerts sometimes omit Case Score Breakdown. In that layout,
        # add the best available JSON user/computer value beneath DIP.
        $stellarNotesTemplate = "Summary:`r`n`r`n" +
                        "SIP: $stellarSourceHost `r`n" +
                        "DIP: $stellarDestinationHost `r`n" +
                        "User: $stellarJsonUser `r`n`r`n" +
                        "Recommendation:`r`n"
    }
    else {
        # Preserve the existing template whenever Case Score Breakdown exists.
        $stellarNotesTemplate = "Summary:`r`n`r`n" +
                        "SIP: $stellarSourceHost `r`n" +
                        "DIP: $stellarDestinationHost `r`n`r`n" +
                        "$stellarCaseScoreBreakdown`r`n`r`n" +
                        "Recommendation:`r`n"
    }

    $stellarSupportingNotesText = $stellarNotesTemplate

    # Send Description to Airtable as one complete URL-prefilled value. Keeping
    # the alert paragraph and template in one value prevents Airtable's rich-text
    # editor from saving later typing at a stale insertion point.
    if ([string]::IsNullOrWhiteSpace($stellarAlertDescription)) {
        $stellarDescriptionText = $stellarSupportingNotesText
    }
    else {
        $stellarDescriptionText = $stellarAlertDescription.TrimEnd() +
                                  "`r`n`r`n" +
                                  $stellarSupportingNotesText
    }

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

    # Paste only the complete JSON after Airtable loads. Issue was already sent
    # in the URL; focus it and leave the caret at the end for the analyst.
    if ($stellarJsonText) {
        [void](Set-AirtableJsonField `
            -JsonText $stellarJsonText `
            -TimeoutSeconds 20
        )
    }

    # Automatically build the same four-provider summary used by the normal
    # IP Lookup tab for every public address found in the submitted case.
    if ($caseIpPublicList.Items.Count -gt 0) {
        & $caseIpRunLookupAction
    }
    else {
        $caseIpStatusLabel.Text = "No public IP addresses were found in this case."
        $caseIpResultsBox.Text = (
            "No public IP addresses were sent to external services.`r`n" +
            "Private and reserved addresses remain listed above."
        )
    }

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
    elseif ($tabControl.SelectedTab -eq $caseIpTab) {
        $mainForm.AcceptButton = $caseIpRunButton
        $caseIpResultsBox.Focus()
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
    $ipSummaryLabel
    $caseIpPrivateLabel
    $caseIpPublicLabel
)

# Calculate the correct bounds once at startup and lock them before display.
Initialize-SocToolsTextBounds `
    -Controls $socToolsTextBoundsControls `
    -Form $mainForm

$mainForm.AcceptButton = $homeOpenButton
[void]$mainForm.ShowDialog()
