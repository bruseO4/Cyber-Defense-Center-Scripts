# SOC Tools - Combined PowerShell GUI
# Home, IP Lookup, and Stellar to Airtable in one window.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# This function runs ONLY in the separate PowerShell worker below. Loading or
# using UI Automation in the GUI process can initialize WPF's DPI awareness
# after the form is already visible, changing its scale and checkbox rendering.
# Paste the requested text fields directly into Airtable instead of relying on
# URL-prefill values. The normal Airtable URL is left unchanged.
# If Airtable has not finished loading, retry until the timeout expires.
function Invoke-AirtableJsonFieldWorker {
    param(
        [Parameter(Mandatory = $true)]
        [string]$JsonText,

        [string]$DescriptionText = "",

        [string]$SupportingNotesText = "",

        [object[]]$AdditionalTextFields = @(),

        [int]$TimeoutSeconds = 20
    )

    $airtableHasAdditionalText = @(
        $AdditionalTextFields |
        Where-Object {
            $null -ne $_ -and
            -not [string]::IsNullOrWhiteSpace([string]$_.Value)
        }
    ).Count -gt 0

    if (
        [string]::IsNullOrWhiteSpace($JsonText) -and
        [string]::IsNullOrWhiteSpace($DescriptionText) -and
        [string]::IsNullOrWhiteSpace($SupportingNotesText) -and
        -not $airtableHasAdditionalText
    ) {
        return $false
    }

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes

    # SendKeys always targets the foreground window. Verify that the exact
    # Airtable Chrome window is still foreground before every keystroke so a
    # delayed retry can never paste into ChatGPT or another application.
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class AirtableNativeWindow
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);

    public static void Click(int x, int y)
    {
        SetCursorPos(x, y);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
}
'@

    $airtableShell = New-Object -ComObject WScript.Shell
    $airtableDeadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)

    $canSendAirtableKeys = {
        param([IntPtr]$ExpectedWindowHandle)

        if ($ExpectedWindowHandle -eq [IntPtr]::Zero) {
            return $false
        }

        return (
            [AirtableNativeWindow]::GetForegroundWindow() -eq
            $ExpectedWindowHandle
        )
    }

    # Keep completion state between retries. A successful field is never
    # selected or pasted again while the remaining controls finish loading.
    $airtableFieldTargets = @(
        [pscustomobject]@{
            Names = @("Json of the event:", "JSON of the event:", "JSON of the event")
            Value = $JsonText
            Order = 100
            Complete = [string]::IsNullOrWhiteSpace($JsonText)
        }
        [pscustomobject]@{
            Names = @("Description:", "Description")
            Value = $DescriptionText
            Order = 103
            Complete = [string]::IsNullOrWhiteSpace($DescriptionText)
        }
        [pscustomobject]@{
            Names = @("Supporting Notes:", "Supporting Notes")
            Value = $SupportingNotesText
            Order = 102
            Complete = [string]::IsNullOrWhiteSpace($SupportingNotesText)
        }
    )

    $airtableNextOrder = 1
    foreach ($airtableAdditionalTextField in @($AdditionalTextFields)) {
        if ($null -eq $airtableAdditionalTextField) {
            continue
        }

        $airtableAdditionalAliases = @(
            $airtableAdditionalTextField.Names |
            ForEach-Object { [string]$_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
        $airtableAdditionalValue = [string]$airtableAdditionalTextField.Value

        if (
            $airtableAdditionalAliases.Count -eq 0 -or
            [string]::IsNullOrWhiteSpace($airtableAdditionalValue)
        ) {
            continue
        }

        $airtableFieldTargets += [pscustomobject]@{
            Names = $airtableAdditionalAliases
            Value = $airtableAdditionalValue
            PressEnter = [bool]$airtableAdditionalTextField.PressEnter
            CloseAfterEnter = [bool]$airtableAdditionalTextField.CloseAfterEnter
            ClickBeforePaste = [bool]$airtableAdditionalTextField.ClickBeforePaste
            FastSelect = [bool]$airtableAdditionalTextField.FastSelect
            Order = $airtableNextOrder
            Complete = $false
        }
        $airtableNextOrder++
    }

    # Fill Description last so later fields cannot interrupt its rich-text
    # paste or replace the clipboard while Chrome is still consuming it.
    $airtableFieldTargets = @($airtableFieldTargets | Sort-Object Order)
    $airtableAnalystLabelBounds = $null
    $airtableFirstAnalystSelected = $false

    $setAirtableClipboardText = {
        param([string]$Value)

        for ($clipboardAttempt = 0; $clipboardAttempt -lt 5; $clipboardAttempt++) {
            try {
                [System.Windows.Forms.Clipboard]::SetText($Value)
            Start-Sleep -Milliseconds 50

                $clipboardReadback = (
                    [System.Windows.Forms.Clipboard]::GetText() -replace "`r`n?", "`n"
                )
                $clipboardExpected = ($Value -replace "`r`n?", "`n")

                if ($clipboardReadback -ceq $clipboardExpected) {
                    return $true
                }
            }
            catch {
                # Chrome or another application may briefly hold the clipboard.
            }

            Start-Sleep -Milliseconds 75
        }

        return $false
    }

    $readAirtableEditorText = {
        param($Element)
        try {
            $pattern = $null
            if ($Element.TryGetCurrentPattern(
                [System.Windows.Automation.ValuePattern]::Pattern, [ref]$pattern
            )) {
                return [pscustomobject]@{ Supported = $true; Text = [string]$pattern.Current.Value }
            }
            $pattern = $null
            if ($Element.TryGetCurrentPattern(
                [System.Windows.Automation.TextPattern]::Pattern, [ref]$pattern
            )) {
                return [pscustomobject]@{ Supported = $true; Text = [string]$pattern.DocumentRange.GetText(-1) }
            }
        }
        catch { } # Replaced editor nodes are resolved again on the next retry.
        return [pscustomobject]@{ Supported = $false; Text = '' }
    }

    $airtableTextMatches = {
        param([string]$Actual, [string]$Expected)
        # Rich-text editors normalize blank lines, trailing newlines, and NBSP.
        # Every word must survive; whitespace layout alone may differ.
        return (($Actual -replace '\s+', ' ').Trim() -ceq
                ($Expected -replace '\s+', ' ').Trim())
    }

    $confirmAirtableDescriptionPaste = {
        param($Element, [string]$ExpectedText, [IntPtr]$WindowHandle)
        $canReadText = $false
        for ($readAttempt = 0; $readAttempt -lt 10; $readAttempt++) {
            if (-not (& $canSendAirtableKeys $WindowHandle)) { return $false }
            $readback = & $readAirtableEditorText $Element
            if ($readback.Supported) {
                $canReadText = $true
                if (& $airtableTextMatches $readback.Text $ExpectedText) { return $true }
            }
            Start-Sleep -Milliseconds 150
        }
        # Prefer editor text over clipboard verification. A mismatch means the
        # complete value must be pasted again, never marked complete silently.
        if ($canReadText) { return $false }

        # Some Chrome rich-text editors do not expose either readback pattern.
        # Copy from the same focused editor; a marker prevents stale clipboard
        # contents from falsely passing when Ctrl+C fails.
        if (-not (& $canSendAirtableKeys $WindowHandle) -or
            -not $Element.Current.HasKeyboardFocus) { return $false }
        $marker = 'SOC description verification ' + [Guid]::NewGuid().ToString()
        if (-not (& $setAirtableClipboardText $marker)) { return $false }
        if (-not (& $canSendAirtableKeys $WindowHandle) -or
            -not $Element.Current.HasKeyboardFocus) { return $false }
        [System.Windows.Forms.SendKeys]::SendWait('^a')
        if (-not (& $canSendAirtableKeys $WindowHandle)) { return $false }
        [System.Windows.Forms.SendKeys]::SendWait('^c')
        for ($copyAttempt = 0; $copyAttempt -lt 6; $copyAttempt++) {
            Start-Sleep -Milliseconds 150
            if (-not (& $canSendAirtableKeys $WindowHandle)) { return $false }
            if (& $airtableTextMatches ([System.Windows.Forms.Clipboard]::GetText()) $ExpectedText) {
                return $true
            }
        }
        return $false
    }

    # Airtable exposes a dropdown's visible label and interactive control as
    # separate accessibility elements. Find the wide focusable control located
    # in the same field container as the matching label.
    $findAirtableDropdownControl = {
        param(
            [System.Windows.Automation.AutomationElement]$LabelElement,
            [string[]]$FieldNames
        )

        try {
            if (
                $LabelElement.Current.IsEnabled -and
                $LabelElement.Current.IsKeyboardFocusable -and
                $LabelElement.Current.ControlType -in @(
                    [System.Windows.Automation.ControlType]::ComboBox,
                    [System.Windows.Automation.ControlType]::Edit,
                    [System.Windows.Automation.ControlType]::Button,
                    [System.Windows.Automation.ControlType]::Custom
                )
            ) {
                return $LabelElement
            }
        }
        catch {
            return $null
        }

        $dropdownTreeWalker = [System.Windows.Automation.TreeWalker]::RawViewWalker
        $dropdownContainer = $LabelElement

        for ($dropdownLevel = 0; $dropdownLevel -lt 6; $dropdownLevel++) {
            try {
                $dropdownContainer = $dropdownTreeWalker.GetParent($dropdownContainer)
            }
            catch {
                return $null
            }

            if ($null -eq $dropdownContainer) {
                break
            }

            try {
                $labelRectangle = $LabelElement.Current.BoundingRectangle
                $dropdownCandidates = $dropdownContainer.FindAll(
                    [System.Windows.Automation.TreeScope]::Descendants,
                    [System.Windows.Automation.Condition]::TrueCondition
                )
                $bestDropdownCandidate = $null
                $bestDropdownScore = [double]::MaxValue

                for (
                    $dropdownCandidateIndex = 0;
                    $dropdownCandidateIndex -lt $dropdownCandidates.Count;
                    $dropdownCandidateIndex++
                ) {
                    $dropdownCandidate = $dropdownCandidates.Item($dropdownCandidateIndex)

                    try {
                        if (
                            -not $dropdownCandidate.Current.IsEnabled -or
                            -not $dropdownCandidate.Current.IsKeyboardFocusable
                        ) {
                            continue
                        }

                        $candidateControlType = $dropdownCandidate.Current.ControlType
                        $allowedControlType = (
                            $candidateControlType -eq [System.Windows.Automation.ControlType]::ComboBox -or
                            $candidateControlType -eq [System.Windows.Automation.ControlType]::Edit -or
                            $candidateControlType -eq [System.Windows.Automation.ControlType]::Button -or
                            $candidateControlType -eq [System.Windows.Automation.ControlType]::Custom
                        )

                        if (-not $allowedControlType) {
                            continue
                        }

                        if (
                            $FieldNames -icontains 'Tenant Name:' -and
                            $dropdownCandidate.Current.Name -match '(?i)Analyst'
                        ) {
                            continue
                        }

                        $candidateRectangle = $dropdownCandidate.Current.BoundingRectangle
                        if ($candidateRectangle.Width -lt 80) {
                            continue
                        }

                        # Prefer an accessibility label relationship when the
                        # browser exposes one. This prevents a nearby Analyst
                        # control from being mistaken for Tenant Name.
                        $candidateLabel = $null
                        try {
                            $candidateLabel = $dropdownCandidate.Current.LabeledBy
                        }
                        catch {
                            # Chromium often omits LabeledBy for custom fields.
                        }

                        if ($null -ne $candidateLabel) {
                            try {
                                if ($FieldNames -inotcontains $candidateLabel.Current.Name) {
                                    continue
                                }
                            }
                            catch {
                                continue
                            }
                        }

                        # Tenant Name sits near Analyst on the form. If its
                        # own control cannot be found directly below its label,
                        # leave it unfilled instead of selecting Analyst.
                        $maximumControlGap = 105
                        if ($FieldNames -icontains 'Tenant Name:') {
                            $maximumControlGap = 75
                        }

                        if (
                            $candidateRectangle.Top -lt ($labelRectangle.Top - 12) -or
                            $candidateRectangle.Top -gt ($labelRectangle.Bottom + $maximumControlGap) -or
                            [Math]::Abs($candidateRectangle.Left - $labelRectangle.Left) -gt 150
                        ) {
                            continue
                        }

                        $controlTypeScore = 3
                        if ($candidateControlType -eq [System.Windows.Automation.ControlType]::ComboBox) {
                            $controlTypeScore = 0
                        }
                        elseif ($candidateControlType -eq [System.Windows.Automation.ControlType]::Edit) {
                            $controlTypeScore = 1
                        }
                        elseif ($candidateControlType -eq [System.Windows.Automation.ControlType]::Button) {
                            $controlTypeScore = 2
                        }

                        # Distance from the matching label decides first;
                        # control type only breaks ties. A farther Analyst
                        # combobox must not beat Tenant's nearby input.
                        $dropdownScore = (
                            $controlTypeScore +
                            ([Math]::Abs($candidateRectangle.Top - $labelRectangle.Bottom) * 10) +
                            [Math]::Abs($candidateRectangle.Left - $labelRectangle.Left)
                        )

                        if ($null -ne $candidateLabel) {
                            $dropdownScore -= 1000
                        }

                        if ($dropdownScore -lt $bestDropdownScore) {
                            $bestDropdownScore = $dropdownScore
                            $bestDropdownCandidate = $dropdownCandidate
                        }
                    }
                    catch {
                        # Ignore accessibility elements replaced during loading.
                    }
                }

                if ($null -ne $bestDropdownCandidate) {
                    return $bestDropdownCandidate
                }
            }
            catch {
                # Try the next parent field container.
            }
        }

        return $null
    }

    $findAnalystAddButton = {
        param(
            [System.Windows.Automation.AutomationElement]$WindowElement,
            $AnalystLabelBounds,
            [int]$SelectedAnalystCount
        )

        # After a linked Analyst record is selected, the next entry starts
        # with the "Add application" button below its record card. Airtable
        # can expose a plus, extra spaces, or a text child in its UIA name;
        # exact-name queries miss those controls.
        try {
            $elements = $WindowElement.FindAll(
                [System.Windows.Automation.TreeScope]::Descendants,
                [System.Windows.Automation.Condition]::TrueCondition
            )
        }
        catch { return $null }
        $bestButton = $null
        $bestScore = [double]::MaxValue

        for ($index = 0; $index -lt $elements.Count; $index++) {
            try {
                $candidate = $elements.Item($index)
                if (-not $candidate.Current.IsEnabled -or $candidate.Current.IsOffscreen) {
                    continue
                }
                $name = [string]$candidate.Current.Name
                if ($name -notmatch '(?i)\bAdd\s+(?:an?\s+)?(?:application|analyst)\b') {
                    continue
                }
                $bounds = $candidate.Current.BoundingRectangle
                if (
                    $bounds.Width -lt 20 -or $bounds.Width -gt 300 -or
                    $bounds.Height -lt 14 -or $bounds.Height -gt 80 -or
                    $bounds.Left -lt ($AnalystLabelBounds.Left - 30) -or
                    $bounds.Left -gt ($AnalystLabelBounds.Left + 180)
                ) { continue }

                $score = [Math]::Abs($bounds.Left - $AnalystLabelBounds.Left)
                if (
                    $candidate.Current.ControlType -ne
                    [System.Windows.Automation.ControlType]::Button
                ) { $score += 100 }
                if ($score -lt $bestScore) {
                    $bestScore = $score
                    $bestButton = $candidate
                }
            }
            catch {
                # Airtable can replace a card during the scan.
            }
        }

        if ($null -ne $bestButton) { return $bestButton }

        # If Chrome omits the button's accessible name, find the visible
        # button in the position shown by the Analyst cards. Each selected
        # record adds a card above the button, so recalculate its row.
        $currentAnalystLabel = $null
        for ($index = 0; $index -lt $elements.Count; $index++) {
            try {
                $candidate = $elements.Item($index)
                if (
                    $candidate.Current.Name -match '(?i)^\s*Analyst\s*:?\s*\*?\s*$' -and
                    -not $candidate.Current.IsOffscreen -and
                    $candidate.Current.BoundingRectangle.Height -lt 45
                ) {
                    $currentAnalystLabel = $candidate.Current.BoundingRectangle
                    break
                }
            }
            catch { }
        }
        if ($null -eq $currentAnalystLabel) { return $null }

        $expectedY = $currentAnalystLabel.Bottom + 142 +
            ([Math]::Max(0, $SelectedAnalystCount - 1) * 110)
        $bestScore = [double]::MaxValue
        for ($index = 0; $index -lt $elements.Count; $index++) {
            try {
                $candidate = $elements.Item($index)
                if (
                    $candidate.Current.ControlType -ne
                        [System.Windows.Automation.ControlType]::Button -or
                    -not $candidate.Current.IsEnabled -or
                    $candidate.Current.IsOffscreen
                ) { continue }
                $bounds = $candidate.Current.BoundingRectangle
                $centerY = $bounds.Top + $bounds.Height / 2
                if (
                    $bounds.Width -lt 95 -or $bounds.Width -gt 240 -or
                    $bounds.Height -lt 25 -or $bounds.Height -gt 55 -or
                    [Math]::Abs($bounds.Left - $currentAnalystLabel.Left) -gt 35 -or
                    [Math]::Abs($centerY - $expectedY) -gt 35
                ) { continue }
                $score = [Math]::Abs($centerY - $expectedY) +
                    [Math]::Abs($bounds.Left - $currentAnalystLabel.Left)
                if ($score -lt $bestScore) {
                    $bestScore = $score
                    $bestButton = $candidate
                }
            }
            catch { }
        }

        return $bestButton
    }

    # Rich-text Description may expose its label separately from an unnamed
    # contenteditable editor. Resolve only within the label's nearest container;
    # never choose the next field or a toolbar button by screen position alone.
    $findAirtableDescriptionEditor = {
        param($LabelElement)

        $walker = [System.Windows.Automation.TreeWalker]::RawViewWalker
        $container = $LabelElement
        for ($level = 0; $level -lt 4; $level++) {
            try {
                $container = $walker.GetParent($container)
                if ($null -eq $container) { break }
                $nodes = $container.FindAll(
                    [System.Windows.Automation.TreeScope]::Descendants,
                    [System.Windows.Automation.Condition]::TrueCondition
                )
                $editors = @()
                $otherField = $false
                for ($nodeIndex = 0; $nodeIndex -lt $nodes.Count; $nodeIndex++) {
                    $node = $nodes.Item($nodeIndex)
                    $name = [string]$node.Current.Name
                    if ($name -match '(?i)^\s*(Supporting Notes|JSON of the event)\s*:?\s*\*?\s*$') {
                        $otherField = $true
                    }
                    if (-not $node.Current.IsEnabled -or -not $node.Current.IsKeyboardFocusable) { continue }
                    if ($node.Current.ControlType -notin @(
                        [System.Windows.Automation.ControlType]::Edit,
                        [System.Windows.Automation.ControlType]::Document
                    )) { continue }
                    # Exclude editor toolbar inputs such as a link URL box.
                    if ($name -and $name -notmatch '(?i)^\s*(Description\s*:?\s*\*?|Rich text editor)\s*$') { continue }
                    $editors += $node
                }
                if ($otherField) { break }
                if ($editors.Count -eq 1) { return $editors[0] }
                if ($editors.Count -gt 1) { break }
            }
            catch { break }
        }
        return $null
    }

    $openAirtableDropdown = {
        param(
            [System.Windows.Automation.AutomationElement]$DropdownControl,
            [IntPtr]$ExpectedWindowHandle
        )

        try {
            if (-not (& $canSendAirtableKeys $ExpectedWindowHandle)) {
                return $false
            }

            $DropdownControl.SetFocus()
            Start-Sleep -Milliseconds 75

            # An editable combobox may already accept pasted text once focused.
            if (
                $DropdownControl.Current.ControlType -eq
                [System.Windows.Automation.ControlType]::Edit
            ) {
                return $true
            }

            $expandPatternObject = $null
            if (
                $DropdownControl.TryGetCurrentPattern(
                    [System.Windows.Automation.ExpandCollapsePattern]::Pattern,
                    [ref]$expandPatternObject
                )
            ) {
                if (
                    $expandPatternObject.Current.ExpandCollapseState -ne
                    [System.Windows.Automation.ExpandCollapseState]::Expanded
                ) {
                    $expandPatternObject.Expand()
                }

                Start-Sleep -Milliseconds 150
                return $true
            }

            $invokePatternObject = $null
            if (
                $DropdownControl.TryGetCurrentPattern(
                    [System.Windows.Automation.InvokePattern]::Pattern,
                    [ref]$invokePatternObject
                )
            ) {
                $invokePatternObject.Invoke()
                Start-Sleep -Milliseconds 150
                return $true
            }

            # Custom Airtable controls without a UIA action normally open with
            # Enter while focused.
            if (-not (& $canSendAirtableKeys $ExpectedWindowHandle)) {
                return $false
            }

            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Start-Sleep -Milliseconds 150
            return $true
        }
        catch {
            return $false
        }
    }

    $airtableActivatedOnce = $false
    $analystAddButtonMisses = 0

    while ([DateTime]::UtcNow -lt $airtableDeadline) {
        # Target only a visible Chrome window whose active tab is Airtable.
        # Never fall back to a generic Chrome window because that could be a
        # different tab or web app.
        $airtableBrowserProcess = @(
            Get-Process -Name "chrome" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.MainWindowHandle -ne 0 -and
                $_.MainWindowTitle -match '(?i)Airtable'
            }
        ) | Select-Object -First 1

        $airtableActivated = $false
        $airtableWindowHandle = [IntPtr]::Zero

        if ($airtableActivatedOnce -and $null -eq $airtableBrowserProcess) {
            return 'AirtableFocusLost'
        }

        if ($null -ne $airtableBrowserProcess) {
            $airtableWindowHandle = [IntPtr]$airtableBrowserProcess.MainWindowHandle
            if ($airtableActivatedOnce) {
                $airtableActivated = & $canSendAirtableKeys $airtableWindowHandle
                if (-not $airtableActivated) { return 'AirtableFocusLost' }
            }
            else {
                $airtableActivated = $airtableShell.AppActivate(
                    $airtableBrowserProcess.Id
                )
                if ($airtableActivated) { $airtableActivatedOnce = $true }
            }
        }

        if ($airtableActivated) {
            Start-Sleep -Milliseconds 350

            try {
                if (-not (& $canSendAirtableKeys $airtableWindowHandle)) {
                    Start-Sleep -Milliseconds 350
                    continue
                }

                # Search only inside the exact Airtable Chrome window.
                $airtableSearchRoot = [System.Windows.Automation.AutomationElement]::FromHandle(
                    $airtableWindowHandle
                )

                if ($null -ne $airtableSearchRoot) {
                    $airtableElements = $airtableSearchRoot.FindAll(
                        [System.Windows.Automation.TreeScope]::Descendants,
                        [System.Windows.Automation.Condition]::TrueCondition
                    )

                    foreach ($airtableFieldTarget in $airtableFieldTargets) {
                        if ($airtableFieldTarget.Complete) {
                            continue
                        }

                        $airtableIsAnalyst = (
                            $airtableFieldTarget.Names -icontains 'Analyst:' -or
                            $airtableFieldTarget.Names -icontains 'Analyst'
                        )

                        # The first selected Analyst becomes a record card.
                        # Click "Add application" below that card to search
                        # for each additional checked Analyst.
                        if ($airtableIsAnalyst -and $airtableFirstAnalystSelected) {
                            $selectedAnalystCount = @(
                                $airtableFieldTargets | Where-Object {
                                    $_.Complete -and $_.Names -icontains 'Analyst:'
                                }
                            ).Count
                            $analystAddButton = & $findAnalystAddButton `
                                $airtableSearchRoot `
                                $airtableAnalystLabelBounds `
                                $selectedAnalystCount
                            if ($null -eq $analystAddButton) {
                                $analystAddButtonMisses++
                                if ($analystAddButtonMisses -ge 3) {
                                    return 'AnalystAddNotFound'
                                }
                                continue
                            }
                            $analystAddButtonMisses = 0

                            if (
                                -not (& $setAirtableClipboardText $airtableFieldTarget.Value) -or
                                -not (& $canSendAirtableKeys $airtableWindowHandle)
                            ) { continue }

                            $addBounds = $analystAddButton.Current.BoundingRectangle
                            [AirtableNativeWindow]::Click(
                                [int][Math]::Floor($addBounds.Left + $addBounds.Width / 2),
                                [int][Math]::Floor($addBounds.Top + $addBounds.Height / 2)
                            )
                            Start-Sleep -Milliseconds 100
                            if (-not (& $canSendAirtableKeys $airtableWindowHandle)) { continue }
                            $analystSearchFocused = $false
                            for ($focusAttempt = 0; $focusAttempt -lt 3; $focusAttempt++) {
                                try {
                                    $focusedControl = [System.Windows.Automation.AutomationElement]::FocusedElement
                                    $focusedType = $focusedControl.Current.ControlType
                                    if (
                                        $focusedType -eq [System.Windows.Automation.ControlType]::Edit -or
                                        $focusedType -eq [System.Windows.Automation.ControlType]::ComboBox -or
                                        (
                                            $focusedType -eq [System.Windows.Automation.ControlType]::Custom -and
                                            $focusedControl.Current.Name -match '(?i)search'
                                        )
                                    ) {
                                        $analystSearchFocused = $true
                                        break
                                    }
                                }
                                catch { }
                                Start-Sleep -Milliseconds 70
                            }
                            if (-not $analystSearchFocused) { return 'AnalystSearchNotFocused' }
                            [System.Windows.Forms.SendKeys]::SendWait("^v")
                            Start-Sleep -Milliseconds 160
                            if (-not (& $canSendAirtableKeys $airtableWindowHandle)) { continue }
                            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                            Start-Sleep -Milliseconds 130
                            $airtableFieldTarget.Complete = $true
                            if (& $canSendAirtableKeys $airtableWindowHandle) {
                                [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
                                Start-Sleep -Milliseconds 130
                            }
                            continue
                        }

                        if ([bool]$airtableFieldTarget.PressEnter) {
                            $airtableDropdownLabel = $null

                            for (
                                $airtableLabelIndex = 0;
                                $airtableLabelIndex -lt $airtableElements.Count;
                                $airtableLabelIndex++
                            ) {
                                $airtableLabelCandidate = $airtableElements.Item(
                                    $airtableLabelIndex
                                )

                                try {
                                    if (
                                        $airtableFieldTarget.Names -icontains
                                        $airtableLabelCandidate.Current.Name -or
                                        (
                                            $airtableFieldTarget.ClickBeforePaste -and
                                            $airtableLabelCandidate.Current.Name -match
                                            '(?i)^\s*Tenant\s+Name\s*:?\s*\*?\s*$'
                                        )
                                    ) {
                                        $airtableDropdownLabel = $airtableLabelCandidate
                                        break
                                    }
                                }
                                catch {
                                    # Airtable may replace a label while loading.
                                }
                            }

                            if ($null -eq $airtableDropdownLabel) {
                                continue
                            }

                            $airtableDropdownControl = $null
                            if (-not $airtableFieldTarget.ClickBeforePaste) {
                                $airtableDropdownControl = & $findAirtableDropdownControl `
                                    $airtableDropdownLabel `
                                    $airtableFieldTarget.Names
                                if ($null -eq $airtableDropdownControl) { continue }
                            }

                            if ($airtableIsAnalyst) {
                                $airtableAnalystLabelBounds = $airtableDropdownLabel.Current.BoundingRectangle
                            }

                            if (
                                -not (& $setAirtableClipboardText $airtableFieldTarget.Value)
                            ) {
                                continue
                            }

                            if ($airtableFieldTarget.ClickBeforePaste) {
                                # The small plus sits 23 pixels to the right
                                # and 33 pixels below the Tenant label in the
                                # form. This needs only the label; the field's
                                # wide control is not exposed consistently.
                                if ($airtableDropdownLabel.Current.IsOffscreen) { continue }
                                $tenantLabelBounds = $airtableDropdownLabel.Current.BoundingRectangle
                                if ($tenantLabelBounds.Width -lt 20) { continue }
                                $tenantPlusX = [int][Math]::Floor($tenantLabelBounds.Left + 23)
                                $tenantPlusY = [int][Math]::Floor($tenantLabelBounds.Bottom + 33)
                                $tenantExpectedX = $tenantPlusX
                                $tenantExpectedY = $tenantPlusY

                                # Prefer the actual small plus button when
                                # Chrome exposes it to accessibility. Its
                                # location is immediately below the left
                                # edge of the Tenant Name label.
                                $plusDistance = [double]::MaxValue
                                for ($plusIndex = 0; $plusIndex -lt $airtableElements.Count; $plusIndex++) {
                                    try {
                                        $plusCandidate = $airtableElements.Item($plusIndex)
                                        if (
                                            -not $plusCandidate.Current.IsEnabled -or
                                            $plusCandidate.Current.IsOffscreen -or
                                            $plusCandidate.Current.ControlType -ne [System.Windows.Automation.ControlType]::Button
                                        ) { continue }
                                        $plusBounds = $plusCandidate.Current.BoundingRectangle
                                        if (
                                            $plusBounds.Width -lt 12 -or $plusBounds.Width -gt 54 -or
                                            $plusBounds.Height -lt 12 -or $plusBounds.Height -gt 54
                                        ) { continue }
                                        $plusName = [string]$plusCandidate.Current.Name
                                        if ($plusName -and $plusName -notmatch '(?i)\+|add|create|new') { continue }
                                        $plusCenterX = $plusBounds.Left + $plusBounds.Width / 2
                                        $plusCenterY = $plusBounds.Top + $plusBounds.Height / 2
                                        $distance = [Math]::Abs($plusCenterX - $tenantExpectedX) +
                                            [Math]::Abs($plusCenterY - $tenantExpectedY)
                                        if (
                                            [Math]::Abs($plusCenterX - $tenantExpectedX) -le 40 -and
                                            [Math]::Abs($plusCenterY - $tenantExpectedY) -le 35 -and
                                            $distance -lt $plusDistance
                                        ) {
                                            $plusDistance = $distance
                                            $tenantPlusX = [int][Math]::Floor($plusCenterX)
                                            $tenantPlusY = [int][Math]::Floor($plusCenterY)
                                        }
                                    }
                                    catch {
                                        # Continue with the label-relative
                                        # position if Airtable replaces a node.
                                    }
                                }
                                if (-not (& $canSendAirtableKeys $airtableWindowHandle)) { continue }
                                [AirtableNativeWindow]::Click($tenantPlusX, $tenantPlusY)
                                Start-Sleep -Milliseconds 300
                            }
                            elseif (
                                -not (& $openAirtableDropdown `
                                    $airtableDropdownControl `
                                    $airtableWindowHandle)
                            ) {
                                continue
                            }

                            # The expanded dropdown's search input now owns the
                            # keyboard focus, matching the successful manual flow.
                            if (-not (& $canSendAirtableKeys $airtableWindowHandle)) {
                                continue
                            }

                            [System.Windows.Forms.SendKeys]::SendWait("^v")
                            # Startup and picker-open waits remain above; once
                            # focused, wait briefly for filtering/selection.
                            $choiceWait = 150
                            if ($airtableFieldTarget.FastSelect) {
                                $choiceWait = 140
                            }
                            elseif ($airtableFieldTarget.CloseAfterEnter) {
                                $choiceWait = 150
                            }
                            Start-Sleep -Milliseconds $choiceWait

                            if (-not (& $canSendAirtableKeys $airtableWindowHandle)) {
                                continue
                            }

                            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                            $acceptWait = 170
                            if ($airtableFieldTarget.FastSelect) {
                                $acceptWait = 150
                            }
                            elseif ($airtableFieldTarget.PressEnter) {
                                # Allow Airtable to resolve and accept the linked
                                # tenant record before moving out of its picker.
                                $acceptWait = 500
                            }
                            elseif ($airtableFieldTarget.CloseAfterEnter) {
                                $acceptWait = 170
                            }
                            Start-Sleep -Milliseconds $acceptWait

                            if (
                                [bool]$airtableFieldTarget.PressEnter -and
                                (& $canSendAirtableKeys $airtableWindowHandle)
                            ) {
                                # Tab commits the selected linked record and
                                # prevents the value remaining only in search.
                                [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
                                Start-Sleep -Milliseconds 250
                            }

                            $airtableFieldTarget.Complete = $true

                            if ($airtableIsAnalyst) {
                                $airtableFirstAnalystSelected = $true
                                if (& $canSendAirtableKeys $airtableWindowHandle) {
                                    [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
                                    Start-Sleep -Milliseconds 130
                                }
                            }

                            # Close any other multi-select dropdown after a
                            # choice has been accepted.
                            if (
                                -not $airtableIsAnalyst -and
                                [bool]$airtableFieldTarget.CloseAfterEnter -and
                                (& $canSendAirtableKeys $airtableWindowHandle)
                            ) {
                                [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
                                Start-Sleep -Milliseconds 100
                            }

                            continue
                        }

                        # Description's accessible name may belong to its label
                        # rather than the rich-text editor. Resolve that editor
                        # before the normal named-input path.
                        $airtableDescriptionEditor = $null
                        if ($airtableFieldTarget.Names -icontains 'Description:') {
                            for ($labelIndex = 0; $labelIndex -lt $airtableElements.Count; $labelIndex++) {
                                try {
                                    $label = $airtableElements.Item($labelIndex)
                                    if ($label.Current.Name -notmatch '(?i)^\s*Description\s*:?\s*\*?\s*$') { continue }
                                    if ($label.Current.IsKeyboardFocusable -and $label.Current.ControlType -in @(
                                        [System.Windows.Automation.ControlType]::Edit,
                                        [System.Windows.Automation.ControlType]::Document
                                    )) {
                                        $airtableDescriptionEditor = $label
                                    }
                                    else {
                                        $airtableDescriptionEditor = & $findAirtableDescriptionEditor $label
                                    }
                                    if ($null -ne $airtableDescriptionEditor) { break }
                                }
                                catch { }
                            }
                            if ($null -eq $airtableDescriptionEditor) { continue }
                        }

                        for (
                            $airtableIndex = 0;
                            $airtableIndex -lt $airtableElements.Count;
                            $airtableIndex++
                        ) {
                            $airtableElement = $airtableElements.Item($airtableIndex)
                            if ($null -ne $airtableDescriptionEditor) {
                                $airtableElement = $airtableDescriptionEditor
                            }

                            try {
                                if (
                                    ($null -eq $airtableDescriptionEditor -and
                                     $airtableFieldTarget.Names -inotcontains $airtableElement.Current.Name) -or
                                    -not $airtableElement.Current.IsEnabled -or
                                    -not $airtableElement.Current.IsKeyboardFocusable
                                ) {
                                    continue
                                }

                                # Verify the clipboard before selecting the field.
                                # This prevents Ctrl+A from firing if Windows has
                                # temporarily locked the clipboard.
                                if (-not (& $setAirtableClipboardText $airtableFieldTarget.Value)) {
                                    break
                                }

                                $scrollPattern = $null
                                if ($airtableElement.TryGetCurrentPattern(
                                    [System.Windows.Automation.ScrollItemPattern]::Pattern,
                                    [ref]$scrollPattern
                                )) {
                                    $scrollPattern.ScrollIntoView()
                                }
                                $airtableElement.SetFocus()
                                Start-Sleep -Milliseconds 75
                                if ($null -ne $airtableDescriptionEditor -and
                                    -not $airtableElement.Current.HasKeyboardFocus) { break }

                                if (-not (& $canSendAirtableKeys $airtableWindowHandle)) {
                                    break
                                }

                                [System.Windows.Forms.SendKeys]::SendWait("^a")
                                Start-Sleep -Milliseconds 30

                                if (-not (& $canSendAirtableKeys $airtableWindowHandle)) {
                                    break
                                }

                                [System.Windows.Forms.SendKeys]::SendWait("^v")
                                Start-Sleep -Milliseconds 100

                                if ($null -ne $airtableDescriptionEditor) {
                                    if (-not (& $confirmAirtableDescriptionPaste `
                                        $airtableElement $airtableFieldTarget.Value $airtableWindowHandle)) {
                                        break
                                    }
                                    # Commit the rich-text edit only after every overview
                                    # and the normal template have been read back in full.
                                    if (-not (& $canSendAirtableKeys $airtableWindowHandle)) { break }
                                    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
                                    Start-Sleep -Milliseconds 400
                                    $committedDescription = & $readAirtableEditorText $airtableElement
                                    if ($committedDescription.Supported -and
                                        -not (& $airtableTextMatches $committedDescription.Text $airtableFieldTarget.Value)) {
                                        break
                                    }
                                }

                                $airtableFieldTarget.Complete = $true
                                break
                            }
                            catch {
                                # Airtable can replace a field while the form loads.
                            }
                        }
                    }

                    $airtableFieldsRemaining = @(
                        $airtableFieldTargets |
                        Where-Object { -not $_.Complete }
                    )

                    if ($airtableFieldsRemaining.Count -eq 0) {
                        return $true
                    }
                }
            }
            catch {
                # Keep retrying while Airtable builds the form and accessibility tree.
            }
        }

        Start-Sleep -Milliseconds 200
    }

    if (
        @($airtableFieldTargets | Where-Object {
            -not $_.Complete -and $_.Names -icontains 'Tenant Name:'
        }).Count -gt 0
    ) {
        return 'TenantNameNotFilled'
    }

    if (@($airtableFieldTargets | Where-Object {
        -not $_.Complete -and $_.Names -icontains 'Description:'
    }).Count -gt 0) {
        return 'DescriptionNotFilled'
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

        [string]$DescriptionText = "",

        [string]$SupportingNotesText = "",

        [object[]]$AdditionalTextFields = @(),

        [ValidateRange(1, 300)]
        [int]$TimeoutSeconds = 20
    )

    $airtableHasAdditionalText = @(
        $AdditionalTextFields |
        Where-Object {
            $null -ne $_ -and
            -not [string]::IsNullOrWhiteSpace([string]$_.Value)
        }
    ).Count -gt 0

    if (
        [string]::IsNullOrWhiteSpace($JsonText) -and
        [string]::IsNullOrWhiteSpace($DescriptionText) -and
        [string]::IsNullOrWhiteSpace($SupportingNotesText) -and
        -not $airtableHasAdditionalText
    ) {
        return $false
    }

    $airtableWorkerProcess = $null
    $airtableWorkerStarted = $false

    try {
        # Preserve the existing clipboard fallback if automatic pasting fails.
        try {
            if (-not [string]::IsNullOrWhiteSpace($JsonText)) {
                [System.Windows.Forms.Clipboard]::SetText($JsonText)
            }
        }
        catch {
            # The worker performs its own clipboard retries before selecting a field.
        }

        $airtableWorkerRequest = @{
            WorkerScript = ${function:Invoke-AirtableJsonFieldWorker}.ToString()
            JsonText = $JsonText
            DescriptionText = $DescriptionText
            SupportingNotesText = $SupportingNotesText
            AdditionalTextFields = @($AdditionalTextFields)
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
    $additionalTextFields = @($request.AdditionalTextFields)
    $pasted = & $worker -JsonText $request.JsonText -DescriptionText $request.DescriptionText -SupportingNotesText $request.SupportingNotesText -AdditionalTextFields $additionalTextFields -TimeoutSeconds $request.TimeoutSeconds
    if ($pasted -eq $true) { exit 0 }
    if ($pasted -eq 'TenantNameNotFilled') { exit 2 }
    if ($pasted -eq 'AnalystAddNotFound') { exit 3 }
    if ($pasted -eq 'AirtableFocusLost') { exit 4 }
    if ($pasted -eq 'AnalystSearchNotFocused') { exit 5 }
    if ($pasted -eq 'DescriptionNotFilled') { exit 6 }
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

        if ($airtableWorkerProcess.ExitCode -eq 2) {
            return 'TenantNameNotFilled'
        }
        if ($airtableWorkerProcess.ExitCode -eq 3) {
            return 'AnalystAddNotFound'
        }
        if ($airtableWorkerProcess.ExitCode -eq 4) {
            return 'AirtableFocusLost'
        }
        if ($airtableWorkerProcess.ExitCode -eq 5) {
            return 'AnalystSearchNotFocused'
        }

        if ($airtableWorkerProcess.ExitCode -eq 6) {
            return 'DescriptionNotFilled'
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

# Find the open BlackSwan/Stellar tab before building the Airtable prefill.
# Chrome does not expose a dependable list of tab URLs to normal PowerShell,
# so this checks each tab through the address bar. If it finds BlackSwan, that
# tab remains selected. If it does not, it returns to the tab that was active
# when the search began and returns an empty value.
function Get-BlackSwanChromeUrl {
    param(
        [int]$MaxTabsToCheck = 30
    )

    $chromeShell = New-Object -ComObject WScript.Shell

    if (-not $chromeShell.AppActivate("Google Chrome")) {
        return ""
    }

    Start-Sleep -Milliseconds 150

    $tabsMoved = 0

    for ($tabIndex = 0; $tabIndex -lt $MaxTabsToCheck; $tabIndex++) {
        [System.Windows.Forms.SendKeys]::SendWait("^l")
        Start-Sleep -Milliseconds 75
        [System.Windows.Forms.SendKeys]::SendWait("^c")
        Start-Sleep -Milliseconds 100

        $candidateUrl = [System.Windows.Forms.Clipboard]::GetText().Trim()
        $isBlackSwanTab = $false

        try {
            $candidateUri = [System.Uri]$candidateUrl
            $isBlackSwanTab = (
                $candidateUri.Host -ieq "blackswan.stellarcyber.cloud" -or
                $candidateUri.Host -ilike "*.blackswan.stellarcyber.cloud"
            )
        }
        catch {
            # A non-URL clipboard value is not a BlackSwan tab.
        }

        if ($isBlackSwanTab) {
            # Keep the BlackSwan tab open and selected for the analyst.
            return ($candidateUrl -split '\?')[0]
        }

        if ($tabIndex -lt ($MaxTabsToCheck - 1)) {
            [System.Windows.Forms.SendKeys]::SendWait("^{TAB}")
            $tabsMoved++
            Start-Sleep -Milliseconds 125
        }
    }

    # No BlackSwan tab was found. Put Chrome back where it started and leave
    # the Stellar URL blank rather than accidentally prefilling another site.
    for ($restoreIndex = 0; $restoreIndex -lt $tabsMoved; $restoreIndex++) {
        [System.Windows.Forms.SendKeys]::SendWait("^+{TAB}")
        Start-Sleep -Milliseconds 50
    }

    return ""
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
# Read each complete pasted JSON in order. Arrays are treated as separate
# alert records so a later record can never supply the primary SIP or DIP.
function Get-StellarJsonEntries {
    param([string]$Text)
    $remaining = $Text
    $seen = New-Object System.Collections.Generic.HashSet[string]
    while (-not [string]::IsNullOrWhiteSpace($remaining)) {
        $block = Get-JsonFromText -Text $remaining
        if ([string]::IsNullOrWhiteSpace($block)) { break }
        $jsonValues = @($block)
        if ($block.TrimStart().StartsWith('[')) {
            $records = ConvertFrom-Json -InputObject $block -ErrorAction Stop
            $jsonValues = @(
                foreach ($record in $records) {
                    if ($null -ne $record -and $record -is [System.Management.Automation.PSCustomObject]) {
                        ConvertTo-Json -InputObject $record -Depth 100
                    }
                }
            )
        }
        foreach ($json in $jsonValues) {
            if ($seen.Add($json)) {
                [pscustomobject]@{
                    Json = $json
                    Description = Get-StellarAlertDescription -JsonText $json
                }
            }
        }
        $end = $remaining.IndexOf($block, [StringComparison]::Ordinal) + $block.Length
        $remaining = $remaining.Substring($end)
    }
}

# Read only Stellar overview string fields directly from the pasted text.
# A complete description remains usable even when its surrounding event cannot
# be converted (for example, a truncated final record or conflicting key names).
# The event JSON and identity fields still require the existing valid-JSON parser.
function Get-StellarOverviewDescriptions {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return }

    $tokens = [regex]::Matches($Text, '"(?:\\.|[^"\\])*"|[{}\[\]:,]')
    $contexts = New-Object System.Collections.Stack
    $pendingProperty = ''
    for ($index = 0; $index -lt $tokens.Count; $index++) {
        $token = $tokens[$index].Value
        if ($token -eq '{' -or $token -eq '[') {
            $contexts.Push([pscustomobject]@{ Property = $pendingProperty; Kind = $token })
            $pendingProperty = ''
            continue
        }
        if ($token -eq '}' -or $token -eq ']') {
            if ($contexts.Count) { [void]$contexts.Pop() }
            $pendingProperty = ''
            continue
        }
        if ($token -eq ',') { $pendingProperty = ''; continue }
        if (-not $token.StartsWith('"')) { continue }

        # Require an actual colon after a key or before a string value. This
        # avoids treating quoted prose elsewhere in the copied case as JSON.
        $afterToken = $tokens[$index].Index + $tokens[$index].Length
        if ($index + 1 -lt $tokens.Count -and $tokens[$index + 1].Value -eq ':' -and
            $Text.Substring($afterToken, $tokens[$index + 1].Index - $afterToken) -match '^\s*$') {
            try { $pendingProperty = (ConvertFrom-Json -InputObject ('{"value":' + $token + '}') -ErrorAction Stop).value }
            catch { $pendingProperty = '' }
            continue
        }
        if ($index -eq 0 -or $tokens[$index - 1].Value -ne ':') { continue }
        $beforeToken = $tokens[$index - 1].Index + 1
        if ($Text.Substring($beforeToken, $tokens[$index].Index - $beforeToken) -notmatch '^\s*$') { continue }
        $isOverview = $pendingProperty -ieq 'xdr_event.description'
        if ($pendingProperty -ieq 'description' -and $contexts.Count -gt 0 -and
            $contexts.Peek().Property -ieq 'xdr_event' -and $contexts.Peek().Kind -eq '{') {
            $isOverview = $true
        }
        if ($isOverview) {
            try {
                $overview = (ConvertFrom-Json -InputObject ('{"value":' + $token + '}') -ErrorAction Stop).value
                $overview = ($overview -replace "`r?`n", "`r`n").Trim()
                if (-not [string]::IsNullOrWhiteSpace($overview)) { $overview }
            }
            catch { } # Never guess at an incomplete description string.
        }
        $pendingProperty = ''
    }
}

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

$airtableCopyTab = New-Object System.Windows.Forms.TabPage
$airtableCopyTab.Text = "Airtable Click and Paste"
$airtableCopyTab.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)

$caseIpTab = New-Object System.Windows.Forms.TabPage
$caseIpTab.Text = "Case IP Lookups"
$caseIpTab.BackColor = [System.Drawing.Color]::FromArgb(245, 248, 252)

[void]$tabControl.TabPages.Add($homeTab)
[void]$tabControl.TabPages.Add($ipTab)
[void]$tabControl.TabPages.Add($stellarTab)
[void]$tabControl.TabPages.Add($airtableCopyTab)
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
$homeRoleComboBox.SelectedIndex = 2
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

    [PSCustomObject]@{
        Name = "IP-DB"
        Url  = "https://ip-db.com/en/ip/{0}"
    }

    [PSCustomObject]@{
        Name = "Shodan InternetDB"
        Url  = "https://internetdb.shodan.io/{0}"
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
#   IPDB_API_KEY (Shodan InternetDB does not require a key)
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

        [hashtable]$Headers = @{},

        [ValidateRange(1, 60)]
        [int]$TimeoutSeconds = 6
    )

    # Older Windows PowerShell installations may not enable TLS 1.2 by default.
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor
        [Net.SecurityProtocolType]::Tls12

    return Invoke-RestMethod `
        -Uri $Uri `
        -Method Get `
        -Headers $Headers `
        -TimeoutSec $TimeoutSeconds `
        -ErrorAction Stop
}

function Get-IpLookupFailureReason {
    param([System.Exception]$Exception)

    if ($null -ne $Exception -and $null -ne $Exception.Response) {
        try {
            $statusCode = [int]$Exception.Response.StatusCode
            if ($statusCode -gt 0) { return "HTTP $statusCode" }
        }
        catch { }
    }
    if ($null -ne $Exception -and $Exception.ToString() -match '(?i)timed?\s*out|timeout|TaskCanceledException') {
        return 'timed out'
    }
    return 'network or service error'
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

# Extract confirmed addresses and keep ambiguous IPv6-like tokens separate.
# Potential fragments are preserved exactly; never invent missing IPv6 groups.
function Get-IpAddressesFromText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

    $candidateMatches = @()
    $colonMatches = @([regex]::Matches(
        $Text,
        '(?i)(?<![a-z0-9_.:%])[0-9a-f:.%]*:[0-9a-f:.%]*(?![a-z0-9_.:%])'
    ))

    foreach ($match in [regex]::Matches(
        $Text, '(?i)(?<![a-z0-9_.])(?:\d{1,3}\.){3}\d{1,3}(?![a-z0-9_]|\.[a-z0-9_.])'
    )) {
        # Do not extract an embedded IPv4 address as a second, unrelated IP.
        $insideIpv6 = $false
        foreach ($colonMatch in $colonMatches) {
            if (@($colonMatch.Value.ToCharArray() | Where-Object { $_ -eq ':' }).Count -ge 2 -and
                $match.Index -ge $colonMatch.Index -and
                ($match.Index + $match.Length) -le ($colonMatch.Index + $colonMatch.Length)) {
                $insideIpv6 = $true
                break
            }
        }
        if (-not $insideIpv6) {
            $candidateMatches += [PSCustomObject]@{ Index = $match.Index; Value = $match.Value; ColonToken = $false }
        }
    }

    foreach ($match in $colonMatches) {
        $token = $match.Value
        # A MAC joined to dotted data is not an IPv6 fragment. Recover only
        # an unambiguous four-octet suffix; never guess within a longer chain.
        if ($token -match '^(?:[0-9a-fA-F]{2}:){5}[0-9a-fA-F]{2}\.(?<IpSuffix>.*)$') {
            $ipSuffix = $matches['IpSuffix']
            if ($ipSuffix -match '^\d{1,3}(?:\.\d{1,3}){3}$' -and
                -not [string]::IsNullOrWhiteSpace((ConvertTo-NormalizedIpAddress -InputText $ipSuffix))) {
                $candidateMatches += [PSCustomObject]@{
                    Index = $match.Index; Value = $ipSuffix; ColonToken = $false
                }
            }
            continue
        }
        if ($token -notmatch '[0-9a-fA-F]') { continue }
        # Exclude JSON property separators, drive letters, clock times, MACs,
        # and IPv4 host:port suffixes from the potential-address column.
        if ($token.StartsWith(':') -and $match.Index -gt 0) {
            $previousCharacter = $Text[$match.Index - 1]
            if ($previousCharacter -eq ']') { continue }
            if ($previousCharacter -eq '"' -and $match.Index -gt 1 -and
                [string]$Text[$match.Index - 2] -match '[a-zA-Z0-9_.-]') { continue }
        }
        $afterIndex = $match.Index + $match.Length
        if ($token -match '^[a-fA-F]:$' -and $afterIndex -lt $Text.Length -and
            ($Text[$afterIndex] -eq '\' -or $Text[$afterIndex] -eq '/')) { continue }
        if ($token -match '^\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?$') { continue }
        if ($token -match '^(?:[0-9a-fA-F]{2}:){5}[0-9a-fA-F]{2}$') { continue }
        if ($token -match '^\d{1,3}(?:\.\d{1,3}){3}:\d+$') { continue }
        $candidateMatches += [PSCustomObject]@{ Index = $match.Index; Value = $token; ColonToken = $true }
    }

    $recordsByAddress = @{}
    $orderedAddressKeys = @()
    foreach ($candidateMatch in @($candidateMatches | Sort-Object Index)) {
        $normalizedAddress = ConvertTo-NormalizedIpAddress -InputText $candidateMatch.Value
        $parsedAddress = $null
        $isPublicAddress = $false
        $isPotential = [string]::IsNullOrWhiteSpace($normalizedAddress)
        if ($isPotential -and -not $candidateMatch.ColonToken) { continue }

        if (-not $isPotential) {
            [void][System.Net.IPAddress]::TryParse($normalizedAddress, [ref]$parsedAddress)
            $isPublicAddress = Test-IpAddressIsPublic -Address $parsedAddress
            # Very short compressed public values (e.g. ::a, e::) can be
            # fragments in alert text. Require the separate potential action.
            if ($parsedAddress.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6 -and
                $isPublicAddress) {
                $explicitGroups = @(($candidateMatch.Value -split ':') | Where-Object { $_ -ne '' })
                if ($explicitGroups.Count -lt 2) { $isPotential = $true }
            }
        }

        $addressValue = $normalizedAddress
        if ($isPotential) { $addressValue = $candidateMatch.Value }
        $addressKey = ('{0}:{1}' -f $isPotential, $addressValue).ToLowerInvariant()
        $sourceLabel = Get-IpAddressSourceLabel -Text $Text -MatchIndex $candidateMatch.Index
        if ($recordsByAddress.ContainsKey($addressKey)) {
            $existingRecord = $recordsByAddress[$addressKey]
            if ($existingRecord.SourceLabels -notcontains $sourceLabel) {
                $existingRecord.SourceLabels = @($existingRecord.SourceLabels) + $sourceLabel
            }
            continue
        }
        $recordsByAddress[$addressKey] = [PSCustomObject]@{
            Address = $addressValue
            ParsedAddress = $parsedAddress
            IsPublic = [bool]$isPublicAddress
            IsPotential = [bool]$isPotential
            SourceLabels = @($sourceLabel)
        }
        $orderedAddressKeys += $addressKey
    }

    foreach ($addressKey in $orderedAddressKeys) {
        $record = $recordsByAddress[$addressKey]
        $sourceDisplay = @($record.SourceLabels) -join ', '
        [PSCustomObject]@{
            Address = $record.Address
            ParsedAddress = $record.ParsedAddress
            IsPublic = $record.IsPublic
            IsPotential = $record.IsPotential
            SourceLabels = @($record.SourceLabels)
            Source = $sourceDisplay
            Display = "$($record.Address)  [$sourceDisplay]"
        }
    }
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

# Return report text only: a single bullet, with original provider numbers.
function Format-IpSocBullet {
    param([string]$IpAddress, [hashtable]$Data, [string[]]$IpInfoDetails,
          [string]$Location, [string]$Network, [string[]]$Unavailable)

    $parts = New-Object System.Collections.Generic.List[string]
    [void]$parts.Add("- $IpAddress - Public IP")
    $attribution = @($IpInfoDetails | Where-Object { $_ -match '^(Location|Organization|ASN):' } |
        ForEach-Object { $_ -replace '^(Location|Organization):\s*', '' -replace '\s+\([A-Z]{2}\)$', '' })
    $info = $Data['IPinfo']
    foreach ($field in @('org', 'isp', 'company')) {
        $value = $info.$field
        if ($null -ne $value -and $value -isnot [string]) { $value = $value.name }
        $value = ConvertTo-IpLookupSingleLine -Value $value -MaximumLength 0
        if ($field -eq 'org') { $value = $value -replace '^AS\d+\s+', '' }
        if ($value -and $attribution -notcontains $value) { $attribution += $value }
    }
    $flags = @($IpInfoDetails | Where-Object { $_ -like 'Threat/context flags:*' } |
        ForEach-Object { $_ -replace '^Threat/context flags:\s*', '' })
    if ($attribution.Count -eq 0) {
        $attribution = @($Location, $Network) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        if ($attribution.Count) { [void]$parts.Add("Attribution: $($attribution -join ', ')") }
        if ($flags.Count) { [void]$parts.Add("IPinfo: $($flags -join '; ')") }
    }
    else {
        $infoText = "IPinfo: $($attribution -join ', ')"
        if ($flags.Count) { $infoText += "; $($flags -join '; ')" }
        [void]$parts.Add($infoText)
    }

    $positive = $false
    $strong = 0
    $negativeChecks = 0
    $exposureNeedsReview = $false
    $otx = $Data['OTX']
    if ($null -ne $otx) {
        $details = @()
        if (Test-IpLookupValueReported $otx.pulse_info 'count') {
            $count = $otx.pulse_info.count
            $details += "$count related pulses"
            if ([double]$count -gt 0) { $positive = $true }
            elseif ([double]$count -eq 0) { $negativeChecks++ }
        }
        if (Test-IpLookupValueReported $otx 'reputation') {
            $details += "reputation $($otx.reputation)"
            # Preserve nonzero provider reputation without guessing its scale.
            if ([string]$otx.reputation -ne '0') { $exposureNeedsReview = $true }
        }
        $pulses = @($otx.pulse_info.pulses | Where-Object { $null -ne $_ })
        $names = @($pulses | ForEach-Object { $_.name } | Where-Object { $_ } | Select-Object -Unique)
        $tags = @($pulses | ForEach-Object { $_.tags } | Where-Object { $_ } | Select-Object -Unique)
        if ($names.Count) { $details += "pulses: $($names -join '; ')"; $positive = $true }
        if ($tags.Count) { $details += "tags: $($tags -join ', ')" }
        if ($details.Count) { [void]$parts.Add("LevelBlue OTX: $($details -join ', ')") }
    }

    $vt = $Data['VirusTotal']
    if ($null -ne $vt) {
        $details = @()
        foreach ($field in @('malicious', 'suspicious', 'harmless', 'undetected')) {
            if (Test-IpLookupValueReported $vt.last_analysis_stats $field) {
                $details += "$($vt.last_analysis_stats.$field) $field"
            }
        }
        if ((Test-IpLookupValueReported $vt.last_analysis_stats 'malicious') -and [double]$vt.last_analysis_stats.malicious -gt 0) { $positive = $true }
        if ((Test-IpLookupValueReported $vt.last_analysis_stats 'suspicious') -and [double]$vt.last_analysis_stats.suspicious -gt 0) { $positive = $true }
        if ([double]$vt.last_analysis_stats.malicious -ge 5) { $strong++ }
        if ((Test-IpLookupValueReported $vt.last_analysis_stats 'malicious') -and
            (Test-IpLookupValueReported $vt.last_analysis_stats 'suspicious') -and
            [double]$vt.last_analysis_stats.malicious -eq 0 -and [double]$vt.last_analysis_stats.suspicious -eq 0) { $negativeChecks++ }
        if (Test-IpLookupValueReported $vt 'reputation') {
            $details += "community reputation $($vt.reputation)"
            if ([double]$vt.reputation -lt 0) { $positive = $true }
        }
        $votes = @()
        foreach ($field in @('malicious', 'harmless')) {
            if (Test-IpLookupValueReported $vt.total_votes $field) { $votes += "$($vt.total_votes.$field) $field" }
        }
        if ($votes.Count) { $details += "community votes: $($votes -join ', ')" }
        if ([double]$vt.total_votes.malicious -gt 0) { $exposureNeedsReview = $true }
        $engines = @($vt.last_analysis_results.PSObject.Properties | Where-Object {
            $_.Value.category -in @('malicious', 'suspicious')
        } | ForEach-Object {
            if ($_.Value.engine_name) { $_.Value.engine_name } else { $_.Name }
        } | Select-Object -Unique)
        if ($engines.Count) { $details += "flagging engines: $($engines -join ', ')"; $positive = $true }
        if (Test-IpLookupValueReported $vt 'last_analysis_date') {
            $details += "last analysis $(ConvertFrom-IpLookupUnixTime $vt.last_analysis_date)"
        }
        if ($details.Count) { [void]$parts.Add("VirusTotal: $($details -join ', ')") }
    }

    $abuse = $Data['AbuseIPDB']
    if ($null -ne $abuse) {
        $details = @()
        if (Test-IpLookupValueReported $abuse 'abuseConfidenceScore') {
            $details += "$($abuse.abuseConfidenceScore)% abuse confidence"
            if ([double]$abuse.abuseConfidenceScore -gt 0) { $positive = $true }
            if ([double]$abuse.abuseConfidenceScore -ge 75) { $strong++ }
        }
        if (Test-IpLookupValueReported $abuse 'totalReports') {
            $details += "$($abuse.totalReports) reports in 90 days"
            if ([double]$abuse.totalReports -gt 0) { $positive = $true }
        }
        if ((Test-IpLookupValueReported $abuse 'abuseConfidenceScore') -and
            (Test-IpLookupValueReported $abuse 'totalReports') -and
            [double]$abuse.abuseConfidenceScore -eq 0 -and [double]$abuse.totalReports -eq 0) { $negativeChecks++ }
        if (Test-IpLookupValueReported $abuse 'numDistinctUsers') { $details += "$($abuse.numDistinctUsers) distinct reporters" }
        if (Test-IpLookupValueReported $abuse 'lastReportedAt') { $details += "last reported $(ConvertTo-IpLookupUtcTime $abuse.lastReportedAt)" }
        if (Test-IpLookupValueReported $abuse 'isTor') { $details += "Tor exit node $(ConvertTo-IpLookupYesNo $abuse.isTor)" }
        if (Test-IpLookupValueReported $abuse 'isWhitelisted') { $details += "whitelisted $(ConvertTo-IpLookupYesNo $abuse.isWhitelisted)" }
        if ($details.Count) { [void]$parts.Add("AbuseIPDB: $($details -join ', ')") }
    }

    $db = $Data['IP-DB']
    if ($null -ne $db) {
        $details = @()
        if (Test-IpLookupValueReported $db 'abuseConfidenceScore') {
            $score = "$($db.abuseConfidenceScore)/100 abuse score"
            if (Test-IpLookupValueReported $db 'scoreLabel') { $score += " ($($db.scoreLabel))" }
            $details += $score
            if ([double]$db.abuseConfidenceScore -gt 0) { $positive = $true }
            if ([double]$db.abuseConfidenceScore -ge 75) { $strong++ }
        }
        elseif (Test-IpLookupValueReported $db 'scoreLabel') { $details += "provider rating $($db.scoreLabel)" }
        if ([string]$db.scoreLabel -match '^(suspicious|critical)$') { $positive = $true }
        foreach ($field in @('totalReports', 'distinctReporters')) {
            if (Test-IpLookupValueReported $db $field) {
                $label = if ($field -eq 'totalReports') { 'reports' } else { 'distinct reporters' }
                $details += "$($db.$field) $label"
            }
        }
        if ([double]$db.totalReports -gt 0) { $positive = $true }
        if (Test-IpLookupValueReported $db 'lastReportedAt') { $details += "last reported $(ConvertTo-IpLookupUtcTime $db.lastReportedAt)" }
        if (Test-IpLookupValueReported $db 'blacklistSources') {
            $lists = @($db.blacklistSources | Where-Object { $_ } | Select-Object -Unique)
            $details += "$($lists.Count) blocklist matches"
            if ($lists.Count) { $details += ($lists -join ', '); $positive = $true }
        }
        if (Test-IpLookupValueReported $db 'dnsbl') {
            $zones = @($db.dnsbl | Where-Object { $null -ne $_ -and $_.listed -ne $false })
            $details += "$($zones.Count) DNSBL listings"
            if ($zones.Count) {
                $details += (@($zones | ForEach-Object {
                    if ($_.label -and $_.label -ne $_.zone) { "$($_.zone) ($($_.label))" } else { $_.zone }
                }) -join ', ')
                $positive = $true
            }
        }
        if ((Test-IpLookupValueReported $db 'abuseConfidenceScore') -and
            (Test-IpLookupValueReported $db 'totalReports') -and
            [double]$db.abuseConfidenceScore -eq 0 -and [double]$db.totalReports -eq 0) { $negativeChecks++ }
        if (Test-IpLookupValueReported $db 'networkType') { $details += "network type $($db.networkType)" }
        if ($null -ne $db.liftedListing) {
            $history = @()
            foreach ($field in @('priorScore', 'delistedAt', 'cycles', 'daysSince')) {
                if (Test-IpLookupValueReported $db.liftedListing $field) { $history += "$field=$($db.liftedListing.$field)" }
            }
            if ($history.Count) { $details += "previous listing: $($history -join ', ')"; $exposureNeedsReview = $true }
        }
        if ($details.Count) { [void]$parts.Add("IP-DB: $($details -join ', ')") }
    }

    $shodan = $Data['Shodan']
    if ($null -ne $shodan) {
        $details = @()
        $ports = @($shodan.ports | Where-Object { $null -ne $_ } | Select-Object -Unique)
        $vulns = @($shodan.vulns | Where-Object { $_ } | Select-Object -Unique)
        $products = @($shodan.cpes | Where-Object { $_ } | ForEach-Object {
            $cpe = [string]$_ -split ':'
            if ($cpe.Count -ge 6 -and $cpe[1] -eq '2.3') { $cpe = $cpe[3..5] }
            elseif ($cpe.Count -ge 5 -and $cpe[1] -match '^/[aoh]$') { $cpe = $cpe[2..4] }
            else { $cpe = @([string]$_) }
            (@($cpe | Where-Object { $_ -and $_ -ne '*' -and $_ -ne '-' }) -join ' ')
        } | Select-Object -Unique)
        $tags = @($shodan.tags | Where-Object { $_ } | Select-Object -Unique)
        if ($ports.Count) { $details += "open ports $($ports -join ', ')" }
        if ($products.Count) { $details += "observed products $($products -join ', ')" }
        if ($tags.Count) { $details += "tags $($tags -join ', ')" }
        if ($vulns.Count) { $details += "potential vulnerabilities $($vulns -join ', ') (verify applicability)"; $exposureNeedsReview = $true }
        if (($tags -join ' ') -match '(?i)malware|botnet|command.and.control|(^|\W)c2($|\W)|compromised') { $positive = $true }
        if ($details.Count) { [void]$parts.Add("Shodan: $($details -join '; ')") }
    }
    if ($Unavailable.Count) { [void]$parts.Add("Coverage limited: $($Unavailable -join ', ')") }
    # Keep the existing 5-engine / 75-score high-risk thresholds, and require
    # corroboration before calling the reputation malicious. This is no case verdict.
    if ($strong -ge 2 -and $null -ne $vt -and [double]$vt.last_analysis_stats.malicious -ge 5) {
        $assessment = 'Malicious reputation - corroborating threat intelligence detections'
    }
    elseif ($positive) { $assessment = 'Suspicious - reported threat or abuse indicators require investigation' }
    elseif ($exposureNeedsReview) { $assessment = 'Requires further investigation - reputation history or exposed services need review' }
    elseif ($negativeChecks -ge 2) { $assessment = 'Clean in available results - no identified threat indicators' }
    else { $assessment = 'Requires further investigation - insufficient reputation data' }
    [void]$parts.Add($assessment)
    # Provider strings are untrusted data, not instructions or Markdown.
    return (($parts -join ' - ') -replace '[\r\n]+', ' ' -replace '[\u2013\u2014]', '-' -replace '[`*]', '')
}

function Format-CaseIpSocReport {
    param([string[]]$Descriptions, [string[]]$IpBullets)
    if ($IpBullets.Count -eq 0) { return '' }
    $introduction = 'IP addresses extracted from the supplied alert text were reviewed using available reputation sources.'
    if (@($IpBullets | Where-Object { $_ -match ' - (Suspicious|Malicious reputation) - ' }).Count) {
        $conclusion = 'Reported threat or abuse indicators require correlation with the observed alert activity. The available evidence does not establish a case verdict.'
    }
    elseif (@($IpBullets | Where-Object { $_ -match 'Requires further investigation' }).Count) {
        $conclusion = 'Further investigation is required to resolve the reputation gaps or findings and determine a case verdict.'
    }
    elseif (@($IpBullets | Where-Object { $_ -match 'Clean in available results' }).Count) {
        $conclusion = 'No threat indicators were identified in the available reputation results. Verify the alert activity before determining a case verdict.'
    }
    else { $conclusion = 'Private or reserved IPs have no public reputation assessment. Verify the alert activity before determining a case verdict.' }
    # IP reputation alone cannot establish whether the alert is a TP or FP.
    $ipSeparator = "`r`n`r`n" + ('-' * 72) + "`r`n`r`n"
    return ((@($introduction, ($IpBullets -join $ipSeparator), $conclusion) -join "`r`n`r`n") -replace '[\u2013\u2014]', '-' -replace '[`*]', '')
}

function Enable-IpBulletClickToCopy {
    param([System.Windows.Forms.TextBox]$TextBox)
    $TextBox.Add_MouseClick({
        param($sender, $eventArgs)
        if ($eventArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        $index = $sender.GetCharIndexFromPosition($eventArgs.Location)
        foreach ($match in [regex]::Matches($sender.Text, '(?m)^- (?<Ip>[0-9a-fA-F:.%]+) - ')) {
            $ipGroup = $match.Groups['Ip']
            if ($index -ge $ipGroup.Index -and $index -lt ($ipGroup.Index + $ipGroup.Length)) {
                $ip = ConvertTo-NormalizedIpAddress $ipGroup.Value
                if ($ip) {
                    $sender.Select($ipGroup.Index, $ipGroup.Length)
                    try {
                        [System.Windows.Forms.Clipboard]::SetText($ip)
                    }
                    catch { } # Leave the selected text available for Ctrl+C.
                }
                break
            }
        }
    })
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
    if ($ParsedAddress.IsIPv4MappedToIPv6) {
        $ParsedAddress = $ParsedAddress.MapToIPv4()
        $IpAddress = $ParsedAddress.ToString()
    }

    if (-not (Test-IpAddressIsPublic -Address $ParsedAddress)) {
        $classification = if ($ParsedAddress.AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork -and
            ($ParsedAddress.GetAddressBytes()[0] -eq 10 -or
             ($ParsedAddress.GetAddressBytes()[0] -eq 172 -and $ParsedAddress.GetAddressBytes()[1] -ge 16 -and $ParsedAddress.GetAddressBytes()[1] -le 31) -or
             ($ParsedAddress.GetAddressBytes()[0] -eq 192 -and $ParsedAddress.GetAddressBytes()[1] -eq 168))) { 'Private IP' }
            elseif ($ParsedAddress.AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetworkV6 -and ($ParsedAddress.GetAddressBytes()[0] -band 0xFE) -eq 0xFC) { 'Private IP' }
            else { 'Reserved IP' }
        return "- $IpAddress - $classification - Not publicly routable; no external lookup - Public reputation does not apply"
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
    $ipDbDetails = @(
        "Status: Not checked. Add IPDB_API_KEY to the API key JSON file."
    )
    $shodanDetails = @('Status: Not checked.')

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
        $otx = Invoke-IpLookupJsonRequest -Uri $otxUri -Headers $otxHeaders -TimeoutSeconds 20
        $otxPassiveDnsNames = @()

        try {
            $otxPassiveDnsUri = "https://otx.alienvault.com/api/v1/indicators/$otxIndicatorType/$encodedIp/passive_dns"
            $otxPassiveDns = Invoke-IpLookupJsonRequest -Uri $otxPassiveDnsUri -Headers $otxHeaders -TimeoutSeconds 10
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
        $otxDetails = @("Status: Threat-intelligence request unavailable ($(Get-IpLookupFailureReason -Exception $_.Exception)).")
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

    # IP-DB combines blacklist, DNSBL, and report evidence in one response.
    $ipDbApiKey = Get-IpLookupApiKey -Name "IPDB_API_KEY"
    if (-not [string]::IsNullOrWhiteSpace($ipDbApiKey)) {
        try {
            $ipDbUri = "https://ip-db.com/api/v1/check/$encodedIp"
            $ipDb = Invoke-IpLookupJsonRequest -Uri $ipDbUri -Headers @{ Key = $ipDbApiKey } -TimeoutSeconds 10
            $providerResponses++
            $threatSourcesChecked++

            $ipDbScore = $null
            if (Test-IpLookupValueReported -InputObject $ipDb -PropertyName 'abuseConfidenceScore') {
                $ipDbScore = [int]$ipDb.abuseConfidenceScore
                $threatSourcesWithAssessment++
            }
            $ipDbBlacklists = @(
                @($ipDb.blacklistSources) |
                    ForEach-Object { ConvertTo-IpLookupSingleLine -Value $_ -MaximumLength 90 } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique |
                    Select-Object -First 12
            )
            $ipDbDnsbl = @(
                @($ipDb.dnsbl) |
                    ForEach-Object { ConvertTo-IpLookupSingleLine -Value $_.zone -MaximumLength 70 } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique |
                    Select-Object -First 8
            )
            if ($null -ne $ipDbScore -and $ipDbScore -ge 75) {
                $hasHighThreatSignal = $true
            }
            elseif (($null -ne $ipDbScore -and $ipDbScore -ge 40) -or $ipDbBlacklists.Count -gt 0) {
                $hasPossibleThreatSignal = $true
            }
            if (($null -ne $ipDbScore -and $ipDbScore -ge 40) -or $ipDbBlacklists.Count -gt 0) {
                $socSignals += "IP-DB reports score $ipDbScore/100 and $($ipDbBlacklists.Count) displayed blacklist match(es)."
            }

            $ipDbDetails = @('Status: Reputation data received.')
            if ($null -ne $ipDbScore) {
                $ipDbScoreLabel = ConvertTo-IpLookupSingleLine -Value $ipDb.scoreLabel -MaximumLength 40
                $ipDbDetails += "Threat detection: Abuse score=$ipDbScore/100; label=$ipDbScoreLabel"
            }
            $ipDbLocation = @(
                [string]$ipDb.city
                [string]$ipDb.countryName
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            if ($ipDbLocation.Count -gt 0) { $ipDbDetails += "Location: $($ipDbLocation -join ', ')" }
            if (Test-IpLookupValueReported -InputObject $ipDb -PropertyName 'asn') {
                $ipDbDetails += "ASN: AS$($ipDb.asn)"
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$ipDb.asnOrg)) {
                $ipDbDetails += "Network owner: $($ipDb.asnOrg)"
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$ipDb.isp)) {
                $ipDbDetails += "ISP: $($ipDb.isp)"
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$ipDb.networkType)) {
                $ipDbDetails += "Network type: $($ipDb.networkType)"
            }
            if (Test-IpLookupValueReported -InputObject $ipDb -PropertyName 'totalReports') {
                $ipDbDetails += "Community reports: $($ipDb.totalReports); distinct reporters: $($ipDb.distinctReporters)"
            }
            if ($ipDbBlacklists.Count -gt 0) {
                $ipDbDetails += "Blacklist matches (up to 12): $($ipDbBlacklists -join ', ')"
            }
            if ($ipDbDnsbl.Count -gt 0) {
                $ipDbDetails += "DNSBL listings (up to 8): $($ipDbDnsbl -join ', ')"
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$ipDb.lastReportedAt)) {
                $ipDbDetails += "Last reported: $(ConvertTo-IpLookupUtcTime -Value $ipDb.lastReportedAt)"
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$ipDb.abuseContact.email)) {
                $ipDbDetails += "Abuse contact: $($ipDb.abuseContact.email)"
            }
            if ([string]::IsNullOrWhiteSpace($location) -and $ipDbLocation.Count -gt 0) {
                $location = $ipDbLocation -join ', '
            }
            if ([string]::IsNullOrWhiteSpace($network) -and -not [string]::IsNullOrWhiteSpace([string]$ipDb.asnOrg)) {
                $network = [string]$ipDb.asnOrg
            }
        }
        catch {
            $ipDbDetails = @("Status: Request unavailable ($(Get-IpLookupFailureReason -Exception $_.Exception)).")
        }
    }

    # Shodan InternetDB enriches an IP with recently observed exposed services.
    # This public endpoint does not require an API key.
    try {
        $shodanUri = "https://internetdb.shodan.io/$encodedIp"
        $shodan = Invoke-IpLookupJsonRequest -Uri $shodanUri -TimeoutSeconds 10
        $providerResponses++

        $shodanPorts = @(
            @($shodan.ports) |
                ForEach-Object { [string]$_ } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique |
                Select-Object -First 30
        )
        $shodanVulnerabilities = @(
            @($shodan.vulns) |
                ForEach-Object { ConvertTo-IpLookupSingleLine -Value $_ -MaximumLength 40 } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique |
                Select-Object -First 20
        )
        $shodanHostnames = @(
            @($shodan.hostnames) |
                ForEach-Object { ConvertTo-IpLookupSingleLine -Value $_ -MaximumLength 100 } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique |
                Select-Object -First 10
        )
        $shodanCpes = @(
            @($shodan.cpes) |
                ForEach-Object { ConvertTo-IpLookupSingleLine -Value $_ -MaximumLength 100 } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique |
                Select-Object -First 10
        )
        $shodanTags = @(
            @($shodan.tags) |
                ForEach-Object { ConvertTo-IpLookupSingleLine -Value $_ -MaximumLength 60 } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique |
                Select-Object -First 12
        )

        $shodanDetails = @('Status: InternetDB data received (dataset refreshes weekly).')
        if ($shodanPorts.Count -gt 0) {
            $shodanDetails += "Open ports (up to 30): $($shodanPorts -join ', ')"
        }
        if ($shodanHostnames.Count -gt 0) {
            $shodanDetails += "Hostnames (up to 10): $($shodanHostnames -join ', ')"
        }
        if ($shodanCpes.Count -gt 0) {
            $shodanDetails += "CPEs / detected products (up to 10): $($shodanCpes -join ', ')"
        }
        if ($shodanTags.Count -gt 0) {
            $shodanDetails += "Tags (up to 12): $($shodanTags -join ', ')"
        }
        if ($shodanVulnerabilities.Count -gt 0) {
            $shodanDetails += "Potential vulnerabilities (up to 20): $($shodanVulnerabilities -join ', ')"
        }
        if ($shodanPorts.Count -gt 0 -or $shodanVulnerabilities.Count -gt 0) {
            $socSignals += "Shodan InternetDB lists $($shodanPorts.Count) port(s) and $($shodanVulnerabilities.Count) potential CVE(s); verify current exposure and applicability."
        }
    }
    catch {
        $shodanError = $_.Exception
        $shodanStatusCode = $null
        try { $shodanStatusCode = [int]$shodanError.Response.StatusCode } catch { }
        if ($shodanStatusCode -eq 404) {
            $providerResponses++
            $shodanDetails = @('Status: No Shodan InternetDB record found (HTTP 404).')
        }
        else {
            $shodanDetails = @("Status: Request unavailable ($(Get-IpLookupFailureReason -Exception $shodanError)).")
        }
    }

    # Retain only successful response objects. Missing values are never zeros.
    $socProviderData = @{}
    $socUnavailable = @()
    if ($ipInfoDetails[0] -match 'data received') { $socProviderData['IPinfo'] = $ipInfo }
    foreach ($source in @(
        @{ Name = 'OTX'; Details = $otxDetails; Data = $otx }
        @{ Name = 'VirusTotal'; Details = $virusTotalDetails; Data = $virusTotalAttributes }
        @{ Name = 'AbuseIPDB'; Details = $abuseIpDbDetails; Data = $abuseData }
        @{ Name = 'IP-DB'; Details = $ipDbDetails; Data = $ipDb }
    )) {
        if ($source.Details[0] -match 'data received') { $socProviderData[$source.Name] = $source.Data }
        else {
            $state = if ($source.Details[0] -match 'Not checked') { 'not configured' } else { 'unavailable' }
            $socUnavailable += "$($source.Name) $state"
        }
    }
    if ($shodanDetails[0] -match 'data received') { $socProviderData['Shodan'] = $shodan }
    elseif ($shodanDetails[0] -notmatch 'HTTP 404') { $socUnavailable += 'Shodan unavailable' }
    return Format-IpSocBullet -IpAddress $IpAddress -Data $socProviderData `
        -IpInfoDetails $ipInfoDetails -Location $location -Network $network -Unavailable $socUnavailable
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

$ipSummaryBox = New-Object System.Windows.Forms.TextBox
$ipSummaryBox.Location = New-Object System.Drawing.Point(400,50)
$ipSummaryBox.Size = New-Object System.Drawing.Size(290,460)
$ipSummaryBox.Multiline = $true
Enable-IpBulletClickToCopy -TextBox $ipSummaryBox
$ipSummaryBox.ScrollBars = "Vertical"
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

# Lists the main and supplemental summary sources side by side.
$ipSummarySourcesTip = New-Object System.Windows.Forms.Label
$ipSummarySourcesTip.Text = "TAB 2 - MAIN SOURCES:`n1. IPinfo`n2. LevelBlue OTX`n3. VirusTotal`n4. AbuseIPDB"
$ipSummarySourcesTip.AutoSize = $false
$ipSummarySourcesTip.Size = New-Object System.Drawing.Size(215,82)
$ipSummarySourcesTip.Location = New-Object System.Drawing.Point(20,552)
$ipSummarySourcesTip.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)

$ipSummarySupplementalTip = New-Object System.Windows.Forms.Label
$ipSummarySupplementalTip.Text = "SUPPLEMENTAL SOURCES:`n1. IP-DB`n2. Shodan InternetDB`nOutside the main source set; queried in Tabs 2 and 5."
$ipSummarySupplementalTip.AutoSize = $false
$ipSummarySupplementalTip.Size = New-Object System.Drawing.Size(230,88)
$ipSummarySupplementalTip.Location = New-Object System.Drawing.Point(250,548)
$ipSummarySupplementalTip.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)


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
$ipSummarySupplementalTip.ForeColor = [System.Drawing.Color]::FromArgb(75,85,95)

# Threat/location results styling
$ipSummaryLabel.ForeColor = [System.Drawing.Color]::FromArgb(35,95,160)
$ipSummaryBox.BackColor = [System.Drawing.Color]::White
$ipSummaryBox.ForeColor = [System.Drawing.Color]::FromArgb(35,35,35)
$ipSummaryBox.BorderStyle = "FixedSingle"
$ipSummaryBox.Font = New-Object System.Drawing.Font("Consolas", 9)
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
            "- $ipIp - Public IP - Requires further investigation - Reputation lookup unavailable"
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
$ipTab.Controls.Add($ipSummarySupplementalTip)


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
$stellarLabel.Text = "Paste Stellar alert text and additional JSONs below:"
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
$stellarTextBox.AcceptsReturn = $true
$stellarTextBox.ScrollBars = "Vertical"
$stellarTab.Controls.Add($stellarTextBox)

# ============================================================
# CASE IP LOOKUPS TAB
# ============================================================

$caseIpInstructions = New-Object System.Windows.Forms.Label
$caseIpInstructions.Text = "Paste complete JSON:"
$caseIpInstructions.AutoSize = $false
$caseIpInstructions.Size = New-Object System.Drawing.Size(280, 24)
$caseIpInstructions.Location = New-Object System.Drawing.Point(20, 20)
$caseIpInstructions.ForeColor = [System.Drawing.Color]::FromArgb(35, 45, 55)
$caseIpTab.Controls.Add($caseIpInstructions)

# This box lets the Case IP Lookups tab work on its own, without
# first using Fill Airtable on the Stellar tab.
$caseIpJsonInputBox = New-Object System.Windows.Forms.TextBox
$caseIpJsonInputBox.Location = New-Object System.Drawing.Point(20, 48)
$caseIpJsonInputBox.Size = New-Object System.Drawing.Size(660, 120)
$caseIpJsonInputBox.Multiline = $true
$caseIpJsonInputBox.ScrollBars = "Vertical"
$caseIpJsonInputBox.BackColor = [System.Drawing.Color]::White
$caseIpJsonInputBox.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$caseIpJsonInputBox.BorderStyle = "FixedSingle"
$caseIpJsonInputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$caseIpTab.Controls.Add($caseIpJsonInputBox)

$caseIpCheckJsonButton = New-Object System.Windows.Forms.Button
$caseIpCheckJsonButton.Text = "Check JSON"
$caseIpCheckJsonButton.Size = New-Object System.Drawing.Size(140, 28)
$caseIpCheckJsonButton.Location = New-Object System.Drawing.Point(540, 16)
$caseIpCheckJsonButton.FlatStyle = "Flat"
$caseIpCheckJsonButton.FlatAppearance.BorderSize = 0
$caseIpCheckJsonButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$caseIpCheckJsonButton.ForeColor = [System.Drawing.Color]::White
$caseIpTab.Controls.Add($caseIpCheckJsonButton)

$caseIpPotentialRunButton = New-Object System.Windows.Forms.Button
$caseIpPotentialRunButton.Text = "Run Potential IPs"
$caseIpPotentialRunButton.Size = New-Object System.Drawing.Size(120, 24)
$caseIpPotentialRunButton.Location = New-Object System.Drawing.Point(410, 18)
$caseIpPotentialRunButton.FlatStyle = "Flat"
$caseIpPotentialRunButton.FlatAppearance.BorderSize = 0
$caseIpPotentialRunButton.BackColor = [System.Drawing.Color]::FromArgb(150, 100, 25)
$caseIpPotentialRunButton.ForeColor = [System.Drawing.Color]::White
$caseIpPotentialRunButton.Enabled = $false
$caseIpTab.Controls.Add($caseIpPotentialRunButton)

$caseIpPrivateLabel = New-Object System.Windows.Forms.Label
$caseIpPrivateLabel.Text = "Private / Reserved (0)"
$caseIpPrivateLabel.AutoSize = $true
$caseIpPrivateLabel.Location = New-Object System.Drawing.Point(20, 178)
$caseIpPrivateLabel.ForeColor = [System.Drawing.Color]::FromArgb(190, 90, 35)
$caseIpTab.Controls.Add($caseIpPrivateLabel)

$caseIpPublicLabel = New-Object System.Windows.Forms.Label
$caseIpPublicLabel.Text = "Public (0)"
$caseIpPublicLabel.AutoSize = $true
$caseIpPublicLabel.Location = New-Object System.Drawing.Point(410, 178)
$caseIpPublicLabel.ForeColor = [System.Drawing.Color]::FromArgb(35, 95, 160)
$caseIpTab.Controls.Add($caseIpPublicLabel)

$caseIpPrivateList = New-Object System.Windows.Forms.ListBox
$caseIpPrivateList.Size = New-Object System.Drawing.Size(210, 60)
$caseIpPrivateList.Location = New-Object System.Drawing.Point(20, 200)
$caseIpPrivateList.HorizontalScrollbar = $true
$caseIpPrivateList.BackColor = [System.Drawing.Color]::White
$caseIpPrivateList.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$caseIpTab.Controls.Add($caseIpPrivateList)

$caseIpPotentialLabel = New-Object System.Windows.Forms.Label
$caseIpPotentialLabel.Text = "Potential IP Addresses (0)"
$caseIpPotentialLabel.AutoSize = $true
$caseIpPotentialLabel.Location = New-Object System.Drawing.Point(245, 178)
$caseIpPotentialLabel.ForeColor = [System.Drawing.Color]::FromArgb(150, 100, 25)
$caseIpTab.Controls.Add($caseIpPotentialLabel)

$caseIpPotentialList = New-Object System.Windows.Forms.ListBox
$caseIpPotentialList.Size = New-Object System.Drawing.Size(150, 60)
$caseIpPotentialList.Location = New-Object System.Drawing.Point(245, 200)
$caseIpPotentialList.HorizontalScrollbar = $true
$caseIpPotentialList.BackColor = [System.Drawing.Color]::White
$caseIpPotentialList.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$caseIpTab.Controls.Add($caseIpPotentialList)

$caseIpPublicList = New-Object System.Windows.Forms.ListBox
$caseIpPublicList.Size = New-Object System.Drawing.Size(270, 120)
$caseIpPublicList.Location = New-Object System.Drawing.Point(410, 200)
$caseIpPublicList.HorizontalScrollbar = $true
$caseIpPublicList.BackColor = [System.Drawing.Color]::White
$caseIpPublicList.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$caseIpTab.Controls.Add($caseIpPublicList)

$caseIpRunButton = New-Object System.Windows.Forms.Button
$caseIpRunButton.Text = "Look Up All Public IPs"
$caseIpRunButton.Size = New-Object System.Drawing.Size(165, 26)
$caseIpRunButton.Location = New-Object System.Drawing.Point(20, 352)
$caseIpRunButton.FlatStyle = "Flat"
$caseIpRunButton.FlatAppearance.BorderSize = 0
$caseIpRunButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$caseIpRunButton.ForeColor = [System.Drawing.Color]::White
$caseIpRunButton.Enabled = $false
$caseIpTab.Controls.Add($caseIpRunButton)

$caseIpCopyButton = New-Object System.Windows.Forms.Button
$caseIpCopyButton.Text = "Copy Results"
$caseIpCopyButton.Size = New-Object System.Drawing.Size(100, 26)
$caseIpCopyButton.Location = New-Object System.Drawing.Point(195, 352)
$caseIpCopyButton.FlatStyle = "Flat"
$caseIpCopyButton.FlatAppearance.BorderSize = 0
$caseIpCopyButton.BackColor = [System.Drawing.Color]::FromArgb(90, 105, 120)
$caseIpCopyButton.ForeColor = [System.Drawing.Color]::White
$caseIpCopyButton.Enabled = $false
$caseIpTab.Controls.Add($caseIpCopyButton)

$caseIpClearButton = New-Object System.Windows.Forms.Button
$caseIpClearButton.Text = 'Clear IP Lookup'
$caseIpClearButton.Size = New-Object System.Drawing.Size(120, 26)
$caseIpClearButton.Location = New-Object System.Drawing.Point(305, 352)
$caseIpClearButton.FlatStyle = 'Flat'
$caseIpClearButton.FlatAppearance.BorderSize = 0
$caseIpClearButton.BackColor = [System.Drawing.Color]::FromArgb(150, 65, 55)
$caseIpClearButton.ForeColor = [System.Drawing.Color]::White
$caseIpTab.Controls.Add($caseIpClearButton)

$caseIpStatusLabel = New-Object System.Windows.Forms.Label
$caseIpStatusLabel.Text = "Paste JSON and select Check JSON."
$caseIpStatusLabel.AutoSize = $false
$caseIpStatusLabel.Size = New-Object System.Drawing.Size(245, 26)
$caseIpStatusLabel.Location = New-Object System.Drawing.Point(435, 352)
$caseIpStatusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$caseIpStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 95)
$caseIpTab.Controls.Add($caseIpStatusLabel)

$caseIpResultsBox = New-Object System.Windows.Forms.TextBox
$caseIpResultsBox.Location = New-Object System.Drawing.Point(20, 390)
$caseIpResultsBox.Size = New-Object System.Drawing.Size(660, 230)
$caseIpResultsBox.Multiline = $true
Enable-IpBulletClickToCopy -TextBox $caseIpResultsBox
$caseIpResultsBox.ScrollBars = "Vertical"
$caseIpResultsBox.BackColor = [System.Drawing.Color]::White
$caseIpResultsBox.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$caseIpResultsBox.BorderStyle = "FixedSingle"
$caseIpResultsBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$caseIpResultsBox.Text = "Tab 5 - Case IP Lookups checks the main sources (IPinfo, LevelBlue OTX, VirusTotal, AbuseIPDB) and supplemental sources (IP-DB and Shodan InternetDB, outside the four main sources). The supplemental sources are also queried in Tab 2 - IP Lookup."
$caseIpTab.Controls.Add($caseIpResultsBox)

$caseIpDescriptionLabel = New-Object System.Windows.Forms.LinkLabel
$caseIpDescriptionLabel.Text = 'Description (click to copy)'
$caseIpDescriptionLabel.AutoSize = $false
$caseIpDescriptionLabel.Size = New-Object System.Drawing.Size(375, 16)
$caseIpDescriptionLabel.Location = New-Object System.Drawing.Point(20, 262)
$caseIpDescriptionLabel.ForeColor = [System.Drawing.Color]::FromArgb(35, 45, 55)
$caseIpDescriptionLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
$caseIpDescriptionLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::HoverUnderline
$caseIpDescriptionLabel.Enabled = $false
$caseIpTab.Controls.Add($caseIpDescriptionLabel)

$caseIpDescriptionPanel = New-Object System.Windows.Forms.Panel
$caseIpDescriptionPanel.Location = New-Object System.Drawing.Point(20, 280)
$caseIpDescriptionPanel.Size = New-Object System.Drawing.Size(375, 64)
$caseIpDescriptionPanel.AutoScroll = $true
$caseIpDescriptionPanel.BackColor = [System.Drawing.Color]::White
$caseIpDescriptionPanel.BorderStyle = 'FixedSingle'
$caseIpDescriptionPanel.Cursor = [System.Windows.Forms.Cursors]::Hand
$caseIpTab.Controls.Add($caseIpDescriptionPanel)

# A wrapping label displays the description; it is not an editable text box.
$caseIpDescriptionBox = New-Object System.Windows.Forms.Label
$caseIpDescriptionBox.Location = New-Object System.Drawing.Point(4, 4)
$caseIpDescriptionBox.AutoSize = $true
$caseIpDescriptionBox.MaximumSize = New-Object System.Drawing.Size(345, 0)
$caseIpDescriptionBox.BackColor = [System.Drawing.Color]::Transparent
$caseIpDescriptionBox.ForeColor = [System.Drawing.Color]::FromArgb(35, 35, 35)
$caseIpDescriptionBox.Cursor = [System.Windows.Forms.Cursors]::Hand
$caseIpDescriptionBox.UseMnemonic = $false
$caseIpDescriptionPanel.Controls.Add($caseIpDescriptionBox)

$caseIpDescriptionHistory = New-Object System.Collections.Generic.List[string]
$caseIpDescriptionSeenJson = New-Object System.Collections.Generic.HashSet[string]
# Retain all overview paragraphs in the Description area above the IP report.
$caseIpReportDescriptions = New-Object System.Collections.Generic.List[string]
$caseIpRefreshDescriptionAction = {
    param([string]$Text)
    # Rebuild from this pasted case so clicking Description copies exactly its
    # overview paragraphs, including repeated text, in their original order.
    $caseIpDescriptionHistory.Clear()
    $caseIpDescriptionSeenJson.Clear()
    $caseIpReportDescriptions.Clear()
    foreach ($caseIpIncomingDescription in @(Get-StellarOverviewDescriptions -Text $Text)) {
        $caseIpDescriptionJson = $caseIpIncomingDescription
        if (-not [string]::IsNullOrWhiteSpace($caseIpIncomingDescription) -and
            -not $caseIpReportDescriptions.Contains($caseIpIncomingDescription.Trim())) {
            [void]$caseIpReportDescriptions.Add($caseIpIncomingDescription.Trim())
        }
        if (-not [string]::IsNullOrWhiteSpace($caseIpIncomingDescription)) {
            [void]$caseIpDescriptionHistory.Add($caseIpIncomingDescription.Trim())
        }
    }
    $caseIpDescriptionPanel.AutoScrollPosition = New-Object System.Drawing.Point(0, 0)
    $caseIpDescriptionBox.Text = $caseIpDescriptionHistory -join "`r`n`r`n"
    $caseIpDescriptionLabel.Text = 'Description (click to copy)'
    $caseIpDescriptionLabel.LinkColor = [System.Drawing.Color]::Blue
    $caseIpDescriptionLabel.Enabled = -not [string]::IsNullOrWhiteSpace($caseIpDescriptionBox.Text)
}

# Descriptions accumulate on Check JSON or when additional Tab 3 JSONs are processed.

# The title, text, and unused space in the panel all support hover and copy.
foreach ($caseIpDescriptionControl in @($caseIpDescriptionLabel, $caseIpDescriptionPanel, $caseIpDescriptionBox)) {
    $caseIpDescriptionControl.Add_MouseEnter({
        if (-not [string]::IsNullOrWhiteSpace($caseIpDescriptionBox.Text)) {
            $caseIpDescriptionLabel.BackColor = [System.Drawing.Color]::FromArgb(220, 238, 255)
            $caseIpDescriptionPanel.BackColor = [System.Drawing.Color]::FromArgb(235, 246, 255)
        }
    })
    $caseIpDescriptionControl.Add_MouseLeave({
        $caseIpDescriptionLabel.BackColor = [System.Drawing.Color]::Transparent
        $caseIpDescriptionPanel.BackColor = [System.Drawing.Color]::White
    })
}

$caseIpCopyDescriptionAction = {
    if (-not [string]::IsNullOrWhiteSpace($caseIpDescriptionBox.Text)) {
        try {
            [System.Windows.Forms.Clipboard]::SetText($caseIpDescriptionBox.Text)
            $caseIpDescriptionLabel.Text = 'Description (copied!)'
            $caseIpDescriptionLabel.LinkColor = [System.Drawing.Color]::FromArgb(25, 125, 65)
            $caseIpStatusLabel.Text = 'Description copied.'
        }
        catch {
            $caseIpStatusLabel.Text = 'Clipboard busy; click the Description area again.'
        }
    }
}
$caseIpDescriptionLabel.Add_LinkClicked($caseIpCopyDescriptionAction)
$caseIpDescriptionPanel.Add_Click($caseIpCopyDescriptionAction)
$caseIpDescriptionBox.Add_Click($caseIpCopyDescriptionAction)


$caseIpAddressByDisplay = @{}
# Keep completed report sections until the analyst explicitly clears them.
$caseIpResultHistory = New-Object System.Collections.Generic.List[string]
# Remember normalized public IPs attempted in this case until Clear IP Lookup.
$caseIpLookedUpAddresses = New-Object System.Collections.Generic.HashSet[string]
$caseIpReportedLocalAddresses = New-Object System.Collections.Generic.HashSet[string]
$caseIpRenderReportAction = {
    $caseIpResultsBox.Text = Format-CaseIpSocReport -Descriptions $caseIpReportDescriptions.ToArray() -IpBullets $caseIpResultHistory.ToArray()
}

$caseIpClearAction = {
    $caseIpResultHistory.Clear()
    $caseIpLookedUpAddresses.Clear()
    $caseIpReportedLocalAddresses.Clear()
    $caseIpReportDescriptions.Clear()
    $caseIpResultsBox.Clear()
    $caseIpDescriptionHistory.Clear()
    $caseIpDescriptionSeenJson.Clear()
    $caseIpDescriptionBox.Text = ''
    $caseIpDescriptionPanel.AutoScrollPosition = New-Object System.Drawing.Point(0, 0)
    $caseIpDescriptionPanel.BackColor = [System.Drawing.Color]::White
    $caseIpDescriptionLabel.Text = 'Description (click to copy)'
    $caseIpDescriptionLabel.LinkColor = [System.Drawing.Color]::Blue
    $caseIpDescriptionLabel.BackColor = [System.Drawing.Color]::Transparent
    $caseIpDescriptionLabel.Enabled = $false
    $caseIpCopyButton.Enabled = $false
    $caseIpStatusLabel.Text = 'Description and IP lookup results cleared.'
}
$caseIpClearButton.Add_Click($caseIpClearAction)

# Clicking a list copies the address or potential fragment that was clicked.
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

$caseIpPotentialList.Add_MouseClick({
    param($listControl, $mouseClickEvent)
    $clickedIndex = $listControl.IndexFromPoint($mouseClickEvent.Location)
    if ($clickedIndex -ge 0) {
        $clickedDisplay = [string]$listControl.Items[$clickedIndex]
        $clickedAddress = [string]$caseIpAddressByDisplay[$clickedDisplay]
        if (-not [string]::IsNullOrWhiteSpace($clickedAddress)) {
            [System.Windows.Forms.Clipboard]::SetText($clickedAddress)
            $caseIpStatusLabel.Text = "Copied potential IP: $clickedAddress"
        }
    }
})

# Rebuild all three lists for Check JSON or Fill Airtable. Potential values
# are excluded from automatic lookups; private/reserved IPs stay local.
$caseIpRefreshAction = {
    param([string]$Text)

    $caseIpPrivateList.BeginUpdate()
    $caseIpPublicList.BeginUpdate()
    $caseIpPotentialList.BeginUpdate()

    try {
        $caseIpPrivateList.Items.Clear()
        $caseIpPublicList.Items.Clear()
        $caseIpPotentialList.Items.Clear()
        $caseIpAddressByDisplay.Clear()

        foreach ($caseIpRecord in @(Get-IpAddressesFromText -Text $Text)) {
            $caseIpDisplay = [string]$caseIpRecord.Display
            $caseIpAddressByDisplay[$caseIpDisplay] = [string]$caseIpRecord.Address

            if ($caseIpRecord.IsPotential) {
                [void]$caseIpPotentialList.Items.Add($caseIpDisplay)
            }
            elseif ($caseIpRecord.IsPublic) {
                [void]$caseIpPublicList.Items.Add($caseIpDisplay)
            }
            else {
                [void]$caseIpPrivateList.Items.Add($caseIpDisplay)
                if ($caseIpReportedLocalAddresses.Add([string]$caseIpRecord.Address)) {
                    [void]$caseIpResultHistory.Add((Get-IpThreatLocationSummary -IpAddress $caseIpRecord.Address -ParsedAddress $caseIpRecord.ParsedAddress -ImportantOnly))
                }
            }
        }
    }
    finally {
        $caseIpPrivateList.EndUpdate()
        $caseIpPublicList.EndUpdate()
        $caseIpPotentialList.EndUpdate()
    }

    if ($caseIpResultHistory.Count -gt 0) { & $caseIpRenderReportAction }
    $caseIpPrivateLabel.Text = "Private / Reserved ($($caseIpPrivateList.Items.Count))"
    $caseIpPublicLabel.Text = "Public ($($caseIpPublicList.Items.Count))"
    $caseIpPotentialLabel.Text = "Potential IP Addresses ($($caseIpPotentialList.Items.Count))"
    $caseIpPotentialRunButton.Enabled = ($caseIpPotentialList.Items.Count -gt 0)
    $caseIpRunButton.Enabled = ($caseIpPublicList.Items.Count -gt 0)
    $caseIpCopyButton.Enabled = ($caseIpResultHistory.Count -gt 0)
    if ($caseIpResultHistory.Count -eq 0) {
        $caseIpResultsBox.Text = "Tab 5 - Case IP Lookups checks the main sources (IPinfo, LevelBlue OTX, VirusTotal, AbuseIPDB) and supplemental sources (IP-DB and Shodan InternetDB, outside the four main sources). The supplemental sources are also queried in Tab 2 - IP Lookup."
    }

    if ([string]::IsNullOrWhiteSpace($Text)) {
        $caseIpStatusLabel.Text = "No case text or JSON has been pasted yet."
    }
    else {
        $caseIpTotalCount = $caseIpPrivateList.Items.Count + $caseIpPublicList.Items.Count
        $caseIpStatusLabel.Text = "$caseIpTotalCount confirmed IP(s); $($caseIpPotentialList.Items.Count) potential."
    }
}

$caseIpRunLookupAction = {
    param([switch]$IncludePotential)

    $caseIpLookupItems = @($caseIpPublicList.Items)
    if ($IncludePotential) { $caseIpLookupItems = @($caseIpPotentialList.Items) }
    if ($caseIpLookupItems.Count -eq 0) {
        return
    }

    # Append new API report sections to previous lookups.
    # Disable Clear during DoEvents so an active run cannot repopulate a reset.
    $caseIpClearButton.Enabled = $false
    $caseIpCheckJsonButton.Enabled = $false
    $caseIpPotentialRunButton.Enabled = $false
    $caseIpJsonInputBox.Enabled = $false
    $caseIpRunButton.Enabled = $false
    $caseIpCopyButton.Enabled = $false
    $mainForm.UseWaitCursor = $true
    $caseIpResultSections = $caseIpResultHistory

    try {
        $caseIpLookupNumber = 0
        $caseIpLookupTotal = $caseIpLookupItems.Count
        $caseIpQueriedCount = 0
        $caseIpDuplicateCount = 0
        $caseIpSkippedCount = 0

        foreach ($caseIpListItem in $caseIpLookupItems) {
            $caseIpLookupNumber++
            $caseIpDisplay = [string]$caseIpListItem
            $caseIpAddress = [string]$caseIpAddressByDisplay[$caseIpDisplay]

            if ([string]::IsNullOrWhiteSpace($caseIpAddress)) {
                $displayAddressPart = ($caseIpDisplay -split '\s+\[', 2)[0]
                $caseIpAddress = ConvertTo-NormalizedIpAddress -InputText $displayAddressPart
            }

            $caseIpStatusLabel.Text = "Looking up $caseIpLookupNumber of $caseIpLookupTotal`: $caseIpAddress"
            [System.Windows.Forms.Application]::DoEvents()

            $caseIpNormalizedAddress = ConvertTo-NormalizedIpAddress -InputText $caseIpAddress
            $caseIpParsedAddress = $null
            if ([string]::IsNullOrWhiteSpace($caseIpNormalizedAddress) -or
                -not [System.Net.IPAddress]::TryParse($caseIpNormalizedAddress, [ref]$caseIpParsedAddress)) {
                $caseIpSkippedCount++
                # An invalid fragment is not an IP and does not belong in the report.
                continue
            }
            elseif (-not (Test-IpAddressIsPublic -Address $caseIpParsedAddress)) {
                $caseIpSkippedCount++
                if (-not $caseIpReportedLocalAddresses.Add($caseIpNormalizedAddress)) { continue }
                $caseIpSummary = Get-IpThreatLocationSummary -IpAddress $caseIpNormalizedAddress -ParsedAddress $caseIpParsedAddress -ImportantOnly
            }
            else {
                # Add before querying so repeat JSONs, button clicks, and even
                # failed attempts cannot query the same public IP until Clear.
                if (-not $caseIpLookedUpAddresses.Add($caseIpNormalizedAddress)) {
                    $caseIpDuplicateCount++
                    continue
                }
                $caseIpQueriedCount++
                try {
                    $caseIpSummary = Get-IpThreatLocationSummary `
                        -IpAddress $caseIpNormalizedAddress `
                        -ParsedAddress $caseIpParsedAddress `
                        -ImportantOnly
                }
                catch {
                    $caseIpSummary = "- $caseIpNormalizedAddress - Public IP - Requires further investigation - Reputation lookup unavailable"
                }
            }

            [void]$caseIpResultSections.Add($caseIpSummary)
            & $caseIpRenderReportAction
            $caseIpResultsBox.SelectionStart = $caseIpResultsBox.Text.Length
            $caseIpResultsBox.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }

        $caseIpResultsBox.SelectionStart = $caseIpResultsBox.Text.Length
        $caseIpResultsBox.ScrollToCaret()
        $caseIpCopyButton.Enabled = ($caseIpResultSections.Count -gt 0)
        if ($IncludePotential) {
            $caseIpStatusLabel.Text = "$caseIpQueriedCount new; $caseIpSkippedCount skipped; $caseIpDuplicateCount checked."
        }
        else {
            $caseIpStatusLabel.Text = "$caseIpQueriedCount new; $caseIpDuplicateCount already checked."
        }
    }
    finally {
        $mainForm.UseWaitCursor = $false
        $caseIpClearButton.Enabled = $true
        $caseIpCopyButton.Enabled = ($caseIpResultHistory.Count -gt 0)
        $caseIpCheckJsonButton.Enabled = $true
        $caseIpJsonInputBox.Enabled = $true
        $caseIpPotentialRunButton.Enabled = ($caseIpPotentialList.Items.Count -gt 0)
        $caseIpRunButton.Enabled = ($caseIpPublicList.Items.Count -gt 0)
    }
}

$caseIpRunButton.Add_Click({
    & $caseIpRunLookupAction
})

$caseIpPotentialRunButton.Add_Click({
    & $caseIpRunLookupAction -IncludePotential
})

$caseIpCheckJsonButton.Add_Click({
    $caseIpPastedJson = $caseIpJsonInputBox.Text
    & $caseIpRefreshDescriptionAction -Text $caseIpPastedJson

    # Reuse the same IP extraction used by Fill Airtable, then immediately run
    # the public-IP summaries so this tab can be used independently.
    & $caseIpRefreshAction -Text $caseIpPastedJson

    if ($caseIpPublicList.Items.Count -gt 0) {
        & $caseIpRunLookupAction
    }
    elseif (-not [string]::IsNullOrWhiteSpace($caseIpPastedJson)) {
        $caseIpStatusLabel.Text = "No public IP addresses were found in the pasted JSON."
        if ($caseIpResultHistory.Count -eq 0) {
            $caseIpResultsBox.Text = (
                "No public IP addresses were sent to external services.`r`n" +
                "Private and reserved addresses remain listed above.`r`n" +
                "Potential IPs are held for the Run Potential IPs button."
            )
        }
    }
})

$caseIpJsonInputBox.Add_KeyDown({
    # Match Tab 3: Backspace clears the whole pasted JSON box at once.
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Back) {
        $caseIpJsonInputBox.Clear()
        $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $caseIpCheckJsonButton.PerformClick()
        $_.SuppressKeyPress = $true
    }
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

$stellarRunCaseIpLookupCheckBox = New-Object System.Windows.Forms.CheckBox
$stellarRunCaseIpLookupCheckBox.Text = "Run Case IP Lookup"
$stellarRunCaseIpLookupCheckBox.Location = New-Object System.Drawing.Point(250, 420)
$stellarRunCaseIpLookupCheckBox.Size = New-Object System.Drawing.Size(220, 25)
$stellarRunCaseIpLookupCheckBox.Checked = $true
$stellarRunCaseIpLookupCheckBox.AutoSize = $false
$stellarRunCaseIpLookupCheckBox.BackColor = [System.Drawing.Color]::FromArgb(32, 34, 37)
$stellarRunCaseIpLookupCheckBox.ForeColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
$stellarTab.Controls.Add($stellarRunCaseIpLookupCheckBox)

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

function Invoke-SocOutlookBodyWorker {
    param([string]$Body, [string]$Marker, [int]$TimeoutSeconds = 25)
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SocOutlookWindow {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint id);
}
'@
    # Only the newly opened compose body contains this per-click marker.
    # Require its body label as well, so a page-wide web document cannot match.
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        try {
            $window = [SocOutlookWindow]::GetForegroundWindow()
            [uint32]$ownerId = 0
            [void][SocOutlookWindow]::GetWindowThreadProcessId($window, [ref]$ownerId)
            $owner = [Diagnostics.Process]::GetProcessById($ownerId)
            if ($owner.ProcessName -notin @('olk', 'OUTLOOK')) {
                Start-Sleep -Milliseconds 250
                continue
            }
            $root = [Windows.Automation.AutomationElement]::FromHandle($window)
            $controls = $root.FindAll(
                [Windows.Automation.TreeScope]::Descendants,
                (New-Object Windows.Automation.OrCondition(
                    (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::Document)),
                    (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty, [Windows.Automation.ControlType]::Edit))
                ))
            )
            $bodyCandidates = @()
            foreach ($control in $controls) {
                if ($control.Current.IsOffscreen -or -not $control.Current.IsEnabled -or
                    $control.Current.Name -notmatch '(?i)^(message body|body of (the )?message|message content|body)(\b|$)') { continue }
                $pattern = $null
                if ($control.TryGetCurrentPattern([Windows.Automation.TextPattern]::Pattern, [ref]$pattern) -and
                    $pattern.DocumentRange.GetText(-1).Contains($Marker)) {
                    $bodyCandidates += $control
                }
            }
            if ($bodyCandidates.Count -ne 1) { Start-Sleep -Milliseconds 250; continue }
            $editor = $bodyCandidates[0]
            $editorId = $editor.GetRuntimeId() -join ','
            $hasEditorFocus = {
                if ([SocOutlookWindow]::GetForegroundWindow() -ne $window) { return $false }
                $focused = [Windows.Automation.AutomationElement]::FocusedElement
                for ($depth = 0; $null -ne $focused -and $depth -lt 50; $depth++) {
                    if (($focused.GetRuntimeId() -join ',') -eq $editorId) { return $true }
                    $focused = [Windows.Automation.TreeWalker]::RawViewWalker.GetParent($focused)
                }
                return $false
            }
            $clipboardReady = $false
            for ($attempt = 0; $attempt -lt 5; $attempt++) {
                try {
                    [Windows.Forms.Clipboard]::SetText($Body)
                    if ([Windows.Forms.Clipboard]::GetText() -ceq $Body) { $clipboardReady = $true; break }
                } catch {}
                Start-Sleep -Milliseconds 100
            }
            if (-not $clipboardReady) { return $false }
            $editor.SetFocus()
            if (-not (& $hasEditorFocus)) { return $false }
            [Windows.Forms.SendKeys]::SendWait('^a')
            if (-not (& $hasEditorFocus) -or [Windows.Forms.Clipboard]::GetText() -cne $Body) { return $false }
            [Windows.Forms.SendKeys]::SendWait('^v')
            # Check what the editor exposes after pasting. Never retry a paste.
            $expected = ($Body -replace '\s+', ' ').Trim()
            for ($verifyAttempt = 0; $verifyAttempt -lt 10; $verifyAttempt++) {
                Start-Sleep -Milliseconds 200
                $pattern = $editor.GetCurrentPattern([Windows.Automation.TextPattern]::Pattern)
                $actual = ($pattern.DocumentRange.GetText(-1) -replace '\s+', ' ').Trim()
                if ($actual.Contains($expected)) { return $true }
            }
            return $false
        }
        catch {
            # A stale accessibility node can disappear while Outlook loads.
            # After keystrokes begin, never target another editor or retry.
            if ($null -ne $editor) { return $false }
            Start-Sleep -Milliseconds 250
        }
    }
    return $false
}

function Set-SocOutlookBody {
    param([string]$Body, [string]$Marker)
    $workerProcess = $null
    try {
        $request = [Management.Automation.PSSerializer]::Serialize(@{
            WorkerScript = ${function:Invoke-SocOutlookBodyWorker}.ToString()
            Body = $Body
            Marker = $Marker
        })
        $bootstrap = @'
$ErrorActionPreference = 'Stop'
try {
    $xml = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String([Console]::In.ReadToEnd()))
    $request = [Management.Automation.PSSerializer]::Deserialize($xml)
    $worker = [scriptblock]::Create($request.WorkerScript)
    if ((& $worker -Body $request.Body -Marker $request.Marker) -eq $true) { exit 0 }
} catch {}
exit 1
'@
        $startInfo = New-Object Diagnostics.ProcessStartInfo
        $startInfo.FileName = Join-Path ([Environment]::GetFolderPath('System')) 'WindowsPowerShell\v1.0\powershell.exe'
        $startInfo.Arguments = '-NoLogo -NoProfile -NonInteractive -STA -WindowStyle Hidden -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($bootstrap))
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        $startInfo.RedirectStandardInput = $true
        $workerProcess = New-Object Diagnostics.Process
        $workerProcess.StartInfo = $startInfo
        if (-not $workerProcess.Start()) { return $false }
        # Case data travels over stdin, never in command arguments or source.
        $workerProcess.StandardInput.Write([Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($request)))
        $workerProcess.StandardInput.Close()
        if (-not $workerProcess.WaitForExit(35000)) { return $false }
        return $workerProcess.ExitCode -eq 0
    }
    catch { return $false }
    finally {
        if ($null -ne $workerProcess) {
            try {
                if (-not $workerProcess.HasExited) {
                    $workerProcess.Kill()
                    [void]$workerProcess.WaitForExit(2000)
                }
            } catch {}
            $workerProcess.Dispose()
        }
    }
}


function Open-SocTruePositiveDraft {
    param(
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$Body
    )

    # Use the Windows email handler (including new Outlook), rather than
    # creating a classic Outlook COM instance. This only opens a compose window.
    $encodedSubject = [uri]::EscapeDataString($Subject)
    $mailPrefix = 'mailto:?subject=' + $encodedSubject + '&body='
    $mailUri = $mailPrefix + [uri]::EscapeDataString($Body)
    $needsBodyPaste = $mailUri.Length -gt 1800
    $bodyMarker = 'SOC-DRAFT-' + [guid]::NewGuid().ToString('N')

    if ($needsBodyPaste) {
        # Keep full JSON out of long mailto links. Never silently truncate it.
        $bodyCopied = $false
        for ($copyAttempt = 0; $copyAttempt -lt 5; $copyAttempt++) {
            try {
                [System.Windows.Forms.Clipboard]::SetText($Body)
                if ([System.Windows.Forms.Clipboard]::GetText() -ceq $Body) {
                    $bodyCopied = $true
                    break
                }
            }
            catch {
                # Another application may briefly hold the clipboard.
            }
            Start-Sleep -Milliseconds 100
        }
        if (-not $bodyCopied) {
            throw 'The full email body could not be copied. Close any clipboard dialog and click True Positive Email again.'
        }
        $pasteNotice = "Preparing the complete email draft: $bodyMarker. If automatic filling stops, click the email body and press Ctrl+A, then Ctrl+V to paste the complete draft."
        $mailUri = $mailPrefix + [uri]::EscapeDataString($pasteNotice)
        if ($mailUri.Length -gt 1800) {
            # An unusually long subject also stays out of the launch URI.
            # Show it in a selectable text box below instead of truncating it.
            $mailUri = 'mailto:'
        }
    }

    Start-Process -FilePath $mailUri -ErrorAction Stop

    if ($needsBodyPaste -and $mailUri -ne 'mailto:') {
        $bodyPasted = Set-SocOutlookBody -Body $Body -Marker $bodyMarker
        if ($bodyPasted) {
            [void][System.Windows.Forms.MessageBox]::Show(
                'The complete email body, including Event JSON at the bottom, was filled and checked. Choose your @uscyberdefensecenter.org account in From before sending.',
                'True Positive Draft'
            )
            return
        }
    }

    if ($needsBodyPaste -and $mailUri -eq 'mailto:') {
        $subjectForm = New-Object System.Windows.Forms.Form
        $subjectForm.Text = 'True Positive Email - Subject'
        $subjectForm.Size = New-Object System.Drawing.Size(600, 180)
        $subjectForm.StartPosition = 'CenterScreen'
        $subjectBox = New-Object System.Windows.Forms.TextBox
        $subjectBox.Multiline = $true
        $subjectBox.ReadOnly = $true
        $subjectBox.Dock = 'Fill'
        $subjectBox.Text = $Subject
        $subjectForm.Controls.Add($subjectBox)
        try {
            [void][System.Windows.Forms.MessageBox]::Show(
                "The subject is too long for an email link. First paste the complete body into Outlook. Then copy the subject from the next window.

Choose your @uscyberdefensecenter.org account in From before sending.",
                'True Positive Draft'
            )
            [void]$subjectForm.ShowDialog()
        }
        finally { $subjectForm.Dispose() }
    }
    elseif ($needsBodyPaste) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "The complete email body, including all Event JSON, is copied to your clipboard.

In the new message, click inside the body, press Ctrl+A, then Ctrl+V. The subject is already filled in.

Choose your @uscyberdefensecenter.org account in From before sending.",
            'True Positive Draft - Paste Full Body'
        )
    }
    else {
        [void][System.Windows.Forms.MessageBox]::Show(
            'The email opened for review. Choose your @uscyberdefensecenter.org account in From before sending.',
            'True Positive Draft'
        )
    }
}


function Get-StellarIssueForEmail {
    param([string]$CaseText, [string]$JsonText)

    $issueMatch = [regex]::Match(
        $CaseText,
        '(?ms)^\s*\d+\s*:\s*(?<Issue>.*?)\s*^\s*Run Analysis\s*$'
    )
    if ($issueMatch.Success) {
        $issue = ($issueMatch.Groups['Issue'].Value -replace '\s+', ' ').Trim()
        if ($issue) { return $issue }
    }

    if ($JsonText) {
        try {
            $records = ConvertFrom-Json -InputObject $JsonText -ErrorAction Stop
            foreach ($record in @($records)) {
                $issue = [string]$record.'xdr_event.display_name'
                if (-not $issue -and $null -ne $record.xdr_event) {
                    $issue = [string]$record.xdr_event.display_name
                }
                if ($issue) { return ($issue -replace '\s+', ' ').Trim() }
            }
        }
        catch {
            # Continue with the visible Issue field if the JSON is incomplete.
        }
    }

    $issueMatch = [regex]::Match(
        $CaseText,
        '(?mi)^\s*Issue\s*:?[ \t]*(?:\r?\n[ \t]*)?([^\r\n]+?)\s*$'
    )
    if ($issueMatch.Success) {
        return ($issueMatch.Groups[1].Value -replace '\s+', ' ').Trim()
    }

    return 'True Positive - Stellar Case'
}

$stellarButton = New-Object System.Windows.Forms.Button
$stellarButton.Text = "Fill Airtable"
$stellarButton.Location = New-Object System.Drawing.Point(250, 535)
$stellarButton.Size = New-Object System.Drawing.Size(180, 45)
$stellarTab.Controls.Add($stellarButton)

$stellarFillModeGroup = New-Object System.Windows.Forms.GroupBox
$stellarFillModeGroup.Text = "Fill method"
$stellarFillModeGroup.Location = New-Object System.Drawing.Point(15, 513)
$stellarFillModeGroup.Size = New-Object System.Drawing.Size(230, 122)
$stellarFillModeGroup.ForeColor = [System.Drawing.Color]::White
$stellarTab.Controls.Add($stellarFillModeGroup)

$stellarScriptingMode = New-Object System.Windows.Forms.RadioButton
$stellarScriptingMode.Text = "Scripting (preferred)"
$stellarScriptingMode.Location = New-Object System.Drawing.Point(9, 17)
$stellarScriptingMode.Size = New-Object System.Drawing.Size(210, 22)
$stellarScriptingMode.Checked = $true
$stellarFillModeGroup.Controls.Add($stellarScriptingMode)

$stellarClickPasteMode = New-Object System.Windows.Forms.RadioButton
$stellarClickPasteMode.Text = "Click and paste (From google form)"
$stellarClickPasteMode.Location = New-Object System.Drawing.Point(9, 39)
$stellarClickPasteMode.Size = New-Object System.Drawing.Size(212, 22)
$stellarClickPasteMode.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$stellarFillModeGroup.Controls.Add($stellarClickPasteMode)

$stellarHybridMode = New-Object System.Windows.Forms.RadioButton
$stellarHybridMode.Text = "Hybrid autofill + scripting`r`n(May crash Airtable)"
$stellarHybridMode.Location = New-Object System.Drawing.Point(9, 61)
$stellarHybridMode.Size = New-Object System.Drawing.Size(212, 32)
$stellarHybridMode.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$stellarFillModeGroup.Controls.Add($stellarHybridMode)

$stellarAutofillMode = New-Object System.Windows.Forms.RadioButton
$stellarAutofillMode.Text = "Autofill (may crash airtable)"
$stellarAutofillMode.Location = New-Object System.Drawing.Point(9, 95)
$stellarAutofillMode.Size = New-Object System.Drawing.Size(212, 22)
$stellarAutofillMode.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$stellarFillModeGroup.Controls.Add($stellarAutofillMode)

$stellarTruePositiveButton = New-Object System.Windows.Forms.Button
$stellarTruePositiveButton.Text = "True Positive Email"
$stellarTruePositiveButton.Location = New-Object System.Drawing.Point(455, 535)
$stellarTruePositiveButton.Size = New-Object System.Drawing.Size(180, 45)
$stellarTab.Controls.Add($stellarTruePositiveButton)

$stellarTenantStatusLabel = New-Object System.Windows.Forms.Label
$stellarTenantStatusLabel.Location = New-Object System.Drawing.Point(250, 590)
$stellarTenantStatusLabel.Size = New-Object System.Drawing.Size(420, 42)
$stellarTenantStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(245, 175, 100)
$stellarTenantStatusLabel.Text = ""
$stellarTenantStatusLabel.Visible = $false
$stellarTab.Controls.Add($stellarTenantStatusLabel)

# Click-and-paste uses the current Airtable parser values in a copy list,
# following the manual workflow from the earlier Google Form tab.
$airtableCopyInstructions = New-Object System.Windows.Forms.Label
$airtableCopyInstructions.Text = "Click a row to copy its answer. Paste into Airtable; press Enter for dropdowns."
$airtableCopyInstructions.Location = New-Object System.Drawing.Point(20, 18)
$airtableCopyInstructions.Size = New-Object System.Drawing.Size(505, 45)
$airtableCopyInstructions.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$airtableCopyInstructions.ForeColor = [System.Drawing.Color]::FromArgb(35, 45, 55)
$airtableCopyTab.Controls.Add($airtableCopyInstructions)

$airtableCopyOpenButton = New-Object System.Windows.Forms.Button
$airtableCopyOpenButton.Text = "Open Airtable Form"
$airtableCopyOpenButton.Location = New-Object System.Drawing.Point(535, 15)
$airtableCopyOpenButton.Size = New-Object System.Drawing.Size(145, 30)
$airtableCopyOpenButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$airtableCopyOpenButton.ForeColor = [System.Drawing.Color]::White
$airtableCopyOpenButton.FlatStyle = "Flat"
$airtableCopyOpenButton.FlatAppearance.BorderSize = 0
$airtableCopyOpenButton.Add_Click({ Start-Process "chrome.exe" $stellarAirtableFormUrl })
$airtableCopyTab.Controls.Add($airtableCopyOpenButton)

$airtableCopyGrid = New-Object System.Windows.Forms.DataGridView
$airtableCopyGrid.Location = New-Object System.Drawing.Point(20, 85)
$airtableCopyGrid.Size = New-Object System.Drawing.Size(660, 385)
$airtableCopyGrid.BackgroundColor = [System.Drawing.Color]::White
$airtableCopyGrid.ForeColor = [System.Drawing.Color]::FromArgb(35, 45, 55)
$airtableCopyGrid.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$airtableCopyGrid.ReadOnly = $true
$airtableCopyGrid.AllowUserToAddRows = $false
$airtableCopyGrid.AllowUserToDeleteRows = $false
$airtableCopyGrid.AllowUserToResizeRows = $false
$airtableCopyGrid.RowHeadersVisible = $false
$airtableCopyGrid.MultiSelect = $false
$airtableCopyGrid.SelectionMode = "FullRowSelect"
$airtableCopyGrid.Anchor = "Top, Left, Right"
[void]$airtableCopyGrid.Columns.Add("Field", "Airtable field")
[void]$airtableCopyGrid.Columns.Add("Answer", "Answer (click to copy)")
[void]$airtableCopyGrid.Columns.Add("Action", "Action")
$airtableCopyGrid.Columns["Field"].Width = 175
$airtableCopyGrid.Columns["Answer"].AutoSizeMode = "Fill"
$airtableCopyGrid.Columns["Action"].Width = 105
foreach ($airtableCopyColumn in $airtableCopyGrid.Columns) {
    $airtableCopyColumn.SortMode = "NotSortable"
}
$airtableCopyTab.Controls.Add($airtableCopyGrid)

$airtableCopyPreviewLabel = New-Object System.Windows.Forms.Label
$airtableCopyPreviewLabel.Text = "Full selected answer"
$airtableCopyPreviewLabel.Location = New-Object System.Drawing.Point(20, 480)
$airtableCopyPreviewLabel.Size = New-Object System.Drawing.Size(250, 22)
$airtableCopyPreviewLabel.ForeColor = [System.Drawing.Color]::FromArgb(35, 45, 55)
$airtableCopyTab.Controls.Add($airtableCopyPreviewLabel)

$airtableCopyPreview = New-Object System.Windows.Forms.TextBox
$airtableCopyPreview.Location = New-Object System.Drawing.Point(20, 505)
$airtableCopyPreview.Size = New-Object System.Drawing.Size(660, 72)
$airtableCopyPreview.Multiline = $true
$airtableCopyPreview.ReadOnly = $true
$airtableCopyPreview.ScrollBars = "Both"
$airtableCopyPreview.Font = New-Object System.Drawing.Font("Consolas", 9)
$airtableCopyPreview.Anchor = "Top, Bottom, Left, Right"
$airtableCopyTab.Controls.Add($airtableCopyPreview)

$airtableCopyStatus = New-Object System.Windows.Forms.Label
$airtableCopyStatus.Text = "Choose Click and paste, then press Fill Airtable to build the list."
$airtableCopyStatus.Location = New-Object System.Drawing.Point(20, 588)
$airtableCopyStatus.Size = New-Object System.Drawing.Size(660, 35)
$airtableCopyStatus.ForeColor = [System.Drawing.Color]::FromArgb(55, 95, 135)
$airtableCopyTab.Controls.Add($airtableCopyStatus)

$airtableCopyGrid.Add_CellClick({
    if ($_.RowIndex -lt 0) { return }
    $airtableCopyRow = $airtableCopyGrid.Rows[$_.RowIndex]
    $airtableCopyField = [string]$airtableCopyRow.Cells["Field"].Value
    $airtableCopyValue = [string]$airtableCopyRow.Tag
    $airtableCopyPreview.Text = $airtableCopyValue
    if ([string]::IsNullOrWhiteSpace($airtableCopyValue)) {
        $airtableCopyStatus.Text = "Select or enter $airtableCopyField manually."
        return
    }
    try {
        [System.Windows.Forms.Clipboard]::SetText($airtableCopyValue)
        $airtableCopyStatus.Text = "Copied $airtableCopyField. Paste into Airtable; press Enter if it is a dropdown."
    }
    catch {
        $airtableCopyStatus.Text = "Could not copy ${airtableCopyField}: $($_.Exception.Message)"
    }
})

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

$stellarTruePositiveButton.BackColor = [System.Drawing.Color]::FromArgb(190, 35, 45)
$stellarTruePositiveButton.ForeColor = [System.Drawing.Color]::White
$stellarTruePositiveButton.FlatStyle = "Flat"
$stellarTruePositiveButton.FlatAppearance.BorderSize = 0
$stellarTruePositiveButton.Font = [System.Drawing.Font]::new(
    "Segoe UI", 13, [System.Drawing.FontStyle]::Bold
)
$stellarTruePositiveButton.Cursor = [System.Windows.Forms.Cursors]::Hand

$stellarTruePositiveButton.Add_MouseEnter({
    $stellarTruePositiveButton.BackColor = [System.Drawing.Color]::FromArgb(215, 48, 55)
})

$stellarTruePositiveButton.Add_MouseLeave({
    $stellarTruePositiveButton.BackColor = [System.Drawing.Color]::FromArgb(190, 35, 45)
})

$stellarTruePositiveButton.Add_Click({
    $caseText = $stellarTextBox.Text
    if ([string]::IsNullOrWhiteSpace($caseText)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Paste the Stellar case information before creating the email draft.",
            "True Positive Draft"
        )
        return
    }

    $emailJsonEntries = @(Get-StellarJsonEntries -Text $caseText)
    $caseJson = ''
    if ($emailJsonEntries.Count -gt 0) {
        $caseJson = $emailJsonEntries[0].Json
    }
    $emailIssue = Get-StellarIssueForEmail -CaseText $caseText -JsonText $caseJson
    $emailCaseNumber = ''
    $caseNumberMatch = [regex]::Match(
        $caseText,
        '(?m)^\s*(?<CaseNumber>\d+)\s*:'
    )
    if (-not $caseNumberMatch.Success) {
        $caseNumberMatch = [regex]::Match(
            $caseText,
            '(?im)^\s*Stellar\s*#\s*:?\s*(?<CaseNumber>\d+)\s*$'
        )
    }
    if ($caseNumberMatch.Success) {
        $emailCaseNumber = $caseNumberMatch.Groups['CaseNumber'].Value.Trim()
    }
    $emailSubject = $emailIssue
    if ($emailCaseNumber) {
        $emailSubject = "$emailCaseNumber`: $emailIssue"
    }

    $caseDescription = @(Get-StellarOverviewDescriptions -Text $caseText) -join "`r`n`r`n"
    if (-not $caseDescription) {
        $caseDescription = '[Add case description here]'
    }

    $emailTenantName = ''
    $tenantMatch = [regex]::Match(
        $caseText,
        '(?im)^[\t ]*Tenant[\t ]+Name[\t ]*:[\t ]*(?<TenantValue>[^\r\n]+?)[\t ]*$'
    )
    if ($tenantMatch.Success) {
        $emailTenantName = $tenantMatch.Groups['TenantValue'].Value.Trim()
    }

    if ([string]::IsNullOrWhiteSpace($emailTenantName) -or $emailTenantName -eq '*') {
        $tenantLines = $caseText -split '\r?\n'
        for ($tenantLineIndex = 0; $tenantLineIndex -lt $tenantLines.Count; $tenantLineIndex++) {
            if ($tenantLines[$tenantLineIndex] -notmatch '(?i)^\s*Tenant\s+Name\s*:?\s*\*?\s*$') {
                continue
            }
            for ($tenantValueIndex = $tenantLineIndex + 1; $tenantValueIndex -lt $tenantLines.Count; $tenantValueIndex++) {
                $tenantCandidate = $tenantLines[$tenantValueIndex].Trim()
                if ([string]::IsNullOrWhiteSpace($tenantCandidate) -or $tenantCandidate -eq '*') {
                    continue
                }
                if ($tenantCandidate -match '^(?i)(Analyst|Stellar Severity|Kill Chain Stage|MITRE Tactic|Detection Type|Escalation Required|Stellar #|Issue|URL|Who|What|When|Where|Source Host|Destination Host)\s*:?\s*$') {
                    break
                }
                $emailTenantName = $tenantCandidate
                break
            }
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($emailTenantName)) {
        $tenantJsonMatch = [regex]::Match(
            $caseText,
            '(?i)"(?:tenant[_ -]?name|tenant)"\s*:\s*"(?<TenantValue>[^"]+)"'
        )
        if ($tenantJsonMatch.Success) {
            $emailTenantName = $tenantJsonMatch.Groups['TenantValue'].Value.Trim()
        }
    }
    if ([string]::IsNullOrWhiteSpace($emailTenantName)) {
        $emailTenantName = '[Add tenant name here]'
    }
    if ($emailTenantName -match '(?i)^SSOC\s*[:\-]\s*(.+)$') {
        $emailTenantName = "SSOC-$($matches[1].Trim())"
    }

    # Prefer the open Stellar case tab; the copied text may also contain its URL.
    $caseUrl = Get-BlackSwanChromeUrl
    if (-not $caseUrl) {
        foreach ($urlMatch in [regex]::Matches($caseText, 'https?://[^\s<>"'']+')) {
            try {
                $candidateUrl = [uri]$urlMatch.Value.TrimEnd(
                    [char[]]@('.', ',', ')', ']')
                )
                if (
                    $candidateUrl.Host -ieq 'blackswan.stellarcyber.cloud' -or
                    $candidateUrl.Host -ilike '*.blackswan.stellarcyber.cloud'
                ) {
                    $caseUrl = $candidateUrl.AbsoluteUri
                    break
                }
            }
            catch {
                # Ignore unrelated or malformed links in the case text.
            }
        }
    }

    if (-not $caseUrl) {
        $caseUrl = '[Paste Stellar case link here]'
    }

    # Keep the original case JSON, formatted for review in the email.
    $formattedJson = $caseJson
    if ($caseJson) {
        try {
            $parsedJson = ConvertFrom-Json -InputObject $caseJson -ErrorAction Stop
            $formattedJson = ConvertTo-Json -InputObject $parsedJson -Depth 100
        }
        catch {
            $formattedJson = $caseJson
        }
    }
    if (-not $formattedJson) {
        $formattedJson = '[No event JSON found in the pasted case]'
    }

    try {
        $stellarEmailLinkLabel = if ($emailCaseNumber) {
            "Stellar Link ($emailCaseNumber):"
        }
        else {
            'Stellar Link:'
        }
        $emailBody = "Tenant Name: $emailTenantName`r`n$stellarEmailLinkLabel $caseUrl`r`n`r`n$caseDescription`r`n`r`nRecommendation:`r`n`r`nEvent JSON:`r`n$formattedJson"
        Open-SocTruePositiveDraft -Subject $emailSubject -Body $emailBody
    }
    catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Outlook could not open a draft: $($_.Exception.Message)",
            "True Positive Draft Error"
        )
    }
})



# ==========================================
# BUTTON ACTION
# ==========================================

$stellarButton.Add_Click({

    $stellarText = $stellarTextBox.Text

    # Start each Airtable fill with a clean Tab 5 IP search and report.
    & $caseIpClearAction

    # Only refresh and query case IPs when the parser-tab checkbox is checked.
    if ($stellarRunCaseIpLookupCheckBox.Checked) {
        & $caseIpRefreshAction -Text $stellarText
        $tabControl.SelectedTab = $caseIpTab
        [System.Windows.Forms.Application]::DoEvents()
    }
    else {
        $caseIpStatusLabel.Text = "Case IP lookup skipped (unchecked on Stellar tab)."
    }

    # -----------------------------
    # JSON DATA
    # -----------------------------

    $stellarJsonText = ""
    $stellarJsonSourceIP = ""
    $stellarJsonDestinationIP = ""
    $stellarJsonHostIP = ""
    $stellarJsonUser = ""
    $stellarAlertDescription = ""

    $stellarJsonEntries = @(Get-StellarJsonEntries -Text $stellarText)
    $stellarPrimaryJsonText = ''
    if ($stellarJsonEntries.Count -gt 0) {
        $stellarPrimaryJsonText = $stellarJsonEntries[0].Json
    }
    # Airtable's event JSON and identity fields always use the first record.
    # Keep all records separately for description overviews and IP lookups.
    $stellarJsonText = $stellarPrimaryJsonText
    $stellarAllJsonText = @($stellarJsonEntries | ForEach-Object { $_.Json }) -join "`r`n`r`n"
    $stellarAlertDescription = @(Get-StellarOverviewDescriptions -Text $stellarText) -join "`r`n`r`n"

    # Display and copy every overview together under Description on Tab 5.
    # Retain the original paste so an unconvertible final event cannot hide its overview.
    $caseIpJsonInputBox.Text = $stellarText
    & $caseIpRefreshDescriptionAction -Text $stellarText
    if ($caseIpResultHistory.Count -gt 0) { & $caseIpRenderReportAction }
    # Source IP
    if ($stellarPrimaryJsonText -match '(?i)"srcip"\s*:\s*"([^"]+)"') {
        $stellarJsonSourceIP = $matches[1].Trim()
    }

    # Destination IP
    if ($stellarPrimaryJsonText -match '(?i)"dstip"\s*:\s*"([^"]+)"') {
        $stellarJsonDestinationIP = $matches[1].Trim()
    }

    # Host IP fallback
    if ($stellarPrimaryJsonText -match '(?i)"hostip"\s*:\s*"([^"]+)"') {
        $stellarJsonHostIP = $matches[1].Trim()
    }

    # User fallback for alerts whose Case Score Breakdown did not load.
    # Prefer engid_name, then use computer_name when engid_name is unavailable.
    if ($stellarPrimaryJsonText -match '(?i)"engid_name"\s*:\s*"([^"]+)"') {
        $stellarJsonUser = $matches[1].Trim()
    }
    elseif ($stellarPrimaryJsonText -match '(?i)"computer_name"\s*:\s*"([^"]+)"') {
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

    # Use only a BlackSwan tab. The helper selects that tab when it finds one;
    # otherwise this stays blank and Airtable receives no Stellar URL.
    $stellarUrlValue = Get-BlackSwanChromeUrl


    
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
    if (-not $stellarIssue -and $stellarPrimaryJsonText) {
        try {
            $stellarIssueJson = ConvertFrom-Json `
                -InputObject $stellarPrimaryJsonText `
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

    # Handle both "Tenant Name: value" and a copied label/value pair on
    # separate lines. The previous broad regex could capture the next field
    # label as the tenant or fail to produce a value, silently omitting the
    # Tenant Name fill action.
    $stellarTenantMatch = [regex]::Match(
        $stellarText,
        '(?im)^[\t ]*Tenant[\t ]+Name[\t ]*:[\t ]*(?<TenantValue>[^\r\n]+?)[\t ]*$'
    )

    if ($stellarTenantMatch.Success) {
        $stellarTenantName = $stellarTenantMatch.Groups['TenantValue'].Value.Trim()
        if ($stellarTenantName -eq '*') { $stellarTenantName = '' }
    }

    if ([string]::IsNullOrWhiteSpace($stellarTenantName)) {
        $stellarTenantLines = $stellarText -split '\r?\n'
        for ($stellarTenantLineIndex = 0; $stellarTenantLineIndex -lt $stellarTenantLines.Count; $stellarTenantLineIndex++) {
            if ($stellarTenantLines[$stellarTenantLineIndex] -notmatch '^[\t ]*Tenant[\t ]+Name[\t ]*:?\s*\*?\s*$') {
                continue
            }

            for (
                $stellarTenantValueIndex = $stellarTenantLineIndex + 1;
                $stellarTenantValueIndex -lt $stellarTenantLines.Count;
                $stellarTenantValueIndex++
            ) {
                $stellarTenantCandidate = $stellarTenantLines[$stellarTenantValueIndex].Trim()
                if (
                    [string]::IsNullOrWhiteSpace($stellarTenantCandidate) -or
                    $stellarTenantCandidate -eq '*'
                ) {
                    continue
                }

                # Stop when the next field label appears; never use it as the
                # tenant name when Stellar omitted the value.
                if ($stellarTenantCandidate -match '^(?i)(Analyst|Stellar Severity|Kill Chain Stage|MITRE Tactic|Detection Type|Escalation Required|Stellar #|Issue|URL|Who|What|When|Where|Source Host|Destination Host)\s*:?\s*$') {
                    break
                }

                $stellarTenantName = $stellarTenantCandidate
                break
            }

            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($stellarTenantName)) {
        $stellarTenantMatch = [regex]::Match(
            $stellarText,
            '(?i)"(?:tenant[_ -]?name|tenant)"\s*:\s*"(?<TenantValue>[^"]+)"'
        )

        if ($stellarTenantMatch.Success) {
            $stellarTenantName = $stellarTenantMatch.Groups['TenantValue'].Value.Trim()
        }
    }

    if ([string]::IsNullOrWhiteSpace($stellarTenantName)) {
        $stellarTenantMatch = [regex]::Match(
            $stellarText,
            '(?mi)^[\t ]*(?<TenantValue>SSOC[\t ]*[:\-][^\r\n]+?)[\t ]*$'
        )

        if ($stellarTenantMatch.Success) {
            $stellarTenantName = $stellarTenantMatch.Groups['TenantValue'].Value.Trim()
        }
    }

    # Stellar may use "SSOC: city" while Airtable uses "SSOC-city".
    if ($stellarTenantName -match '(?i)^SSOC\s*[:\-]\s*(.+)$') {
        $stellarTenantName = "SSOC-$($matches[1].Trim())"
    }
    # Correct this known alias before searching Airtable's linked Tenant field.
    if ($stellarTenantName -match '(?i)^SSOC-H-CEI$') {
        $stellarTenantName = 'SSOC-CEI'
    }

    if ([string]::IsNullOrWhiteSpace($stellarTenantName)) {
        $stellarTenantStatusLabel.Text = "Tenant Name was not found in the pasted case; select it manually."
    }
    else {
        $stellarTenantStatusLabel.Text = "Tenant Name found: $stellarTenantName"
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

    $stellarSourceHost = "None"

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

    if ([string]::IsNullOrWhiteSpace($stellarSourceHost)) {
        $stellarSourceHost = "None"
    }

    # -----------------------------
    # DESTINATION HOST
    # -----------------------------

    $stellarDestinationHost = "None"

    if ($stellarJsonDestinationIP) {
        $stellarDestinationHost = $stellarJsonDestinationIP
    }
    elseif ($stellarSourceHost) {
        $stellarDestinationHost = $stellarSourceHost
    }

    if ([string]::IsNullOrWhiteSpace($stellarDestinationHost)) {
        $stellarDestinationHost = "None"
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

    # Keep Description as one complete rich-text value. The worker pastes it
    # into Airtable after the form loads, preventing stale editor insertion.
    if ([string]::IsNullOrWhiteSpace($stellarAlertDescription)) {
        $stellarDescriptionText = $stellarSupportingNotesText
    }
    else {
        $stellarDescriptionText = $stellarAlertDescription.TrimEnd() +
                                  "`r`n`r`n" +
                                  $stellarSupportingNotesText
    }

    # Description only: keep two blank lines after Recommendation for manual
    # analyst notes. Supporting Notes remains unchanged.
    $stellarDescriptionText += "`r`n`r`n"

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
    if ($stellarMitreTactic -ieq "XDR EBA") {
        $stellarMitreTactic = "XDR UBA"
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


    # JSON, Description, and Supporting Notes are scripted after Airtable loads.
    # Keep these long text fields out of the URL to avoid oversized links.
    
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


    

   

    


    # Each click uses the selected mode. The default scripting and manual
    # copy modes keep the original Airtable URL with no query string.
    $stellarFillMode = if ($stellarClickPasteMode.Checked) {
        'copy'
    }
    elseif ($stellarHybridMode.Checked) {
        'hybrid'
    }
    elseif ($stellarAutofillMode.Checked) {
        'autofill'
    }
    else {
        'scripting'
    }
    $stellarFinalUrl = $stellarAirtableFormUrl

    # Hybrid prefills only Tenant Name and Analysts.
    $stellarHybridPrefillKeys = @(
        'prefill_Tenant+Name%3A', 'prefill_Analyst'
    )
    if ($stellarFillMode -in @('autofill', 'hybrid')) {
        $stellarAutofillParameters = @($stellarParameters)
        if ($stellarFillMode -eq 'hybrid') {
            $stellarAutofillParameters = @($stellarParameters | Where-Object {
                $stellarHybridPrefillKeys -contains ($_ -split '=', 2)[0]
            })
        }


        if ($stellarAutofillParameters.Count -gt 0) {
            $stellarFinalUrl += '?' + ($stellarAutofillParameters -join '&')
        }
    }

    $stellarSelectedAnalysts = @(
        $stellarAnalystChecklist.CheckedItems |
        ForEach-Object { $_.ToString() }
    )

    $airtableAdditionalTextFields = @(
        # Start at the top of the form with Tenant Name.
        [pscustomobject]@{
            Names = @(
                "Tenant Name:",
                "Tenant Name",
                "Tenant Name: *",
                "Tenant Name *"
            )
            Value = $stellarTenantName
            PressEnter = $true
            ClickBeforePaste = $true
        }
    )

    # Add each checked Analyst immediately after Tenant Name, in checklist order.
    foreach ($stellarSelectedAnalyst in $stellarSelectedAnalysts) {
        $airtableAdditionalTextFields += [pscustomobject]@{
            Names = @(
                "Analyst:",
                "Analyst",
                "Analyst: *",
                "Analyst *",
                "Analysts",
                "Analysts on Shift"
            )
            Value = $stellarSelectedAnalyst
            PressEnter = $true
            FastSelect = $true
        }
    }

    $airtableAdditionalTextFields += @(
        # Then follow the visible form downward. Analyst Determination is manual.
        [pscustomobject]@{
            Names = @("Stellar Severity", "Stellar Severity *", "Severity")
            Value = $stellarSeverity
            PressEnter = $true
        }
        [pscustomobject]@{
            Names = @("Kill Chain Stage", "Kill Chain Stage *")
            Value = $stellarKillChainStage
            PressEnter = $true
        }
        [pscustomobject]@{
            Names = @("MITRE Tactic", "MITRE Tactic *", "MITRE Technique")
            Value = $stellarMitreTactic
            PressEnter = $true
        }
        [pscustomobject]@{
            Names = @("Detection Type", "Detection Type *")
            Value = $stellarDetectionType
            PressEnter = $true
        }
        [pscustomobject]@{
            Names = @(
                "Escalation Required?",
                "Escalation Required? *",
                "Escalation Required"
            )
            Value = "Mentor Review"
            PressEnter = $true
        }
        [pscustomobject]@{
            Names = @("Stellar #", "Stellar #:", "Stellar # *")
            Value = $stellarStellar
        }
        [pscustomobject]@{
            Names = @("Issue:", "Issue", "Issue: *")
            Value = $stellarIssue
        }
        [pscustomobject]@{
            Names = @("URL:", "URL", "URL: *", "Stellar URL", "Case URL")
            Value = $stellarUrlValue
        }
        [pscustomobject]@{
            Names = @("Who", "Who *", "Who:")
            Value = $stellarWho
        }
        [pscustomobject]@{
            Names = @("What", "What *", "What:")
            Value = $stellarWhat
        }
        [pscustomobject]@{
            Names = @("When", "When *", "When:")
            Value = $stellarWhen
        }
        [pscustomobject]@{
            Names = @("Where", "Where *", "Where:")
            Value = $stellarWhere
        }
        [pscustomobject]@{
            Names = @("Source Host:", "Source Host", "Source Host: *")
            Value = $stellarSourceHost
        }
        [pscustomobject]@{
            Names = @("Destination Host:", "Destination Host", "Destination Host: *")
            Value = $stellarDestinationHost
        }
    )

    if ($stellarFillMode -eq 'copy') {
        $airtableCopyValues = [ordered]@{}
        $airtableCopyValues['Tenant Name:'] = $stellarTenantName
        if ($stellarSelectedAnalysts.Count -gt 0) {
            $airtableCopyValues['Analyst'] = $stellarSelectedAnalysts -join ', '
        }
        else {
            $airtableCopyValues['Analyst'] = ''
        }
        $airtableCopyValues['Stellar Severity'] = $stellarSeverity
        $airtableCopyValues['Kill Chain Stage'] = $stellarKillChainStage
        $airtableCopyValues['MITRE Tactic'] = $stellarMitreTactic
        $airtableCopyValues['Analyst Determination'] = ''
        $airtableCopyValues['Detection Type'] = $stellarDetectionType
        $airtableCopyValues['Escalation Required?'] = 'Mentor Review'
        $airtableCopyValues['Stellar #'] = $stellarStellar
        $airtableCopyValues['Issue:'] = $stellarIssue
        $airtableCopyValues['URL:'] = $stellarUrlValue
        $airtableCopyValues['Who'] = $stellarWho
        $airtableCopyValues['What'] = $stellarWhat
        $airtableCopyValues['When'] = $stellarWhen
        $airtableCopyValues['Where'] = $stellarWhere
        $airtableCopyValues['Source Host:'] = $stellarSourceHost
        $airtableCopyValues['Destination Host:'] = $stellarDestinationHost
        $airtableCopyValues['JSON of the event:'] = $stellarJsonText
        $airtableCopyValues['Description:'] = $stellarDescriptionText
        $airtableCopyValues['Supporting Notes:'] = $stellarSupportingNotesText

        $airtableCopyGrid.Rows.Clear()
        foreach ($airtableCopyField in $airtableCopyValues.Keys) {
            $airtableCopyValue = [string]$airtableCopyValues[$airtableCopyField]
            $airtableCopyAction = if ($airtableCopyField -eq 'Analyst Determination') {
                'Select manually'
            }
            elseif ([string]::IsNullOrWhiteSpace($airtableCopyValue)) {
                'Missing'
            }
            else {
                'Click to copy'
            }
            $airtableCopyExcerpt = ($airtableCopyValue -replace '\s+', ' ').Trim()
            if ($airtableCopyExcerpt.Length -gt 110) {
                $airtableCopyExcerpt = $airtableCopyExcerpt.Substring(0, 107) + '...'
            }
            $airtableCopyRowIndex = $airtableCopyGrid.Rows.Add()
            $airtableCopyRow = $airtableCopyGrid.Rows[$airtableCopyRowIndex]
            $airtableCopyRow.Cells['Field'].Value = $airtableCopyField
            $airtableCopyRow.Cells['Answer'].Value = $airtableCopyExcerpt
            $airtableCopyRow.Cells['Action'].Value = $airtableCopyAction
            $airtableCopyRow.Tag = $airtableCopyValue
        }
        $airtableCopyPreview.Clear()
        $airtableCopyStatus.Text = 'Click each row to copy its answer. Checked analysts are together in one row.'
    }

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

    # Hybrid leaves only Tenant Name and Analysts under URL control,
    # then scripts Stellar Severity onward. Autofill scripts only the final text sections.
    if ($stellarFillMode -in @('scripting', 'autofill', 'hybrid') -and (
        $stellarJsonText -or
        $stellarDescriptionText -or
        $stellarSupportingNotesText -or
        ($stellarFillMode -in @('scripting', 'hybrid') -and $airtableAdditionalTextFields.Count -gt 0)
    )) {
        $airtableWorkerAdditionalFields = if ($stellarFillMode -eq 'scripting') {
            $airtableAdditionalTextFields
        }
        elseif ($stellarFillMode -eq 'hybrid') {
            # Slice at Stellar Severity so scripts follow form order.
            $stellarHybridStart = 0
            while ($stellarHybridStart -lt $airtableAdditionalTextFields.Count -and
                $airtableAdditionalTextFields[$stellarHybridStart].Names -notcontains 'Stellar Severity') {
                $stellarHybridStart++
            }
            @($airtableAdditionalTextFields | Select-Object -Skip $stellarHybridStart)
        }
        else {
            @()
        }
        $airtableFillResult = Set-AirtableJsonField `
            -JsonText $stellarJsonText `
            -DescriptionText $stellarDescriptionText `
            -SupportingNotesText $stellarSupportingNotesText `
            -AdditionalTextFields $airtableWorkerAdditionalFields `
            -TimeoutSeconds 75

        if ($airtableFillResult -eq 'TenantNameNotFilled') {
            $stellarTenantStatusLabel.Text = "Tenant Name was found in the case, but Airtable's field could not be selected."
        }
        elseif ($airtableFillResult -eq 'AnalystAddNotFound') {
            $stellarTenantStatusLabel.Text = "Add application was not found after the first Analyst; remaining names were not added."
        }
        elseif ($airtableFillResult -eq 'AnalystSearchNotFocused') {
            $stellarTenantStatusLabel.Text = "Add application was clicked, but the Analyst search field did not receive focus."
        }
        elseif ($airtableFillResult -eq 'DescriptionNotFilled') {
            $stellarTenantStatusLabel.Text = "Airtable's Description editor could not be filled; use Airtable Copy for the complete description."
        }
        elseif ($airtableFillResult -eq 'AirtableFocusLost') {
            $stellarTenantStatusLabel.Text = "Airtable filling stopped when Chrome lost focus."
        }
    }

    # Automatically build the same six-provider summary only when enabled.
    if ($stellarRunCaseIpLookupCheckBox.Checked) {
        if ($caseIpPublicList.Items.Count -gt 0) {
            & $caseIpRunLookupAction
        }
        else {
            $caseIpStatusLabel.Text = "No public IP addresses were found in this case."
            if ($caseIpResultHistory.Count -eq 0) {
                $caseIpResultsBox.Text = (
                    "No public IP addresses were sent to external services.`r`n" +
                    "Private and reserved addresses remain listed above.`r`n" +
                    "Potential IPs are held for the Run Potential IPs button."
                )
            }
        }
    }

    if ($stellarFillMode -eq 'copy') {
        $tabControl.SelectedTab = $airtableCopyTab
        [void]$mainForm.Activate()
    }
    elseif ($stellarFillMode -eq 'hybrid' -and $airtableFillResult -eq $true) {
        $stellarTenantStatusLabel.Text = 'Hybrid opened Airtable; Tenant and Analysts use autofill; Stellar Severity onward uses script.'
    }
    elseif ($stellarFillMode -eq 'autofill' -and $airtableFillResult -eq $true) {
        $stellarTenantStatusLabel.Text = 'Autofill opened Airtable; JSON, Description, and Supporting Notes are being filled by script.'
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

    # Enter adds a newline. Fill Airtable runs only from its button.

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
        $mainForm.AcceptButton = $null
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
