# Loads the Windows Forms library so PowerShell can create a GUI window
Add-Type -AssemblyName System.Windows.Forms

# Loads drawing tools used for window/button sizes and positions
Add-Type -AssemblyName System.Drawing


# ------------------------------------------------------------
# HISTORY FILE
# ------------------------------------------------------------

# Saves lookup history in the same folder as this PowerShell script
$historyFile = "$PSScriptRoot\IP-Lookup-History.txt"

# Saves which lookup websites are checked or unchecked
$settingsFile = "$PSScriptRoot\IP-Lookup-Settings.json"

# Chrome executable location
$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
# ------------------------------------------------------------
# LOOKUP WEBSITE DEFINITIONS
# ------------------------------------------------------------

# Each website has a display name and URL.
# {0} will later be replaced with the IP address.
$sites = @(
    
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
$sites = $sites | Sort-Object Name

# ------------------------------------------------------------
# MAIN WINDOW
# ------------------------------------------------------------

# Creates the main application window
$form = New-Object System.Windows.Forms.Form

# Text shown at the top of the window
$form.Text = "IP Lookup"

# Sets the width and height of the window
$form.Size = New-Object System.Drawing.Size(420,610)

# Makes the window open in the center of the screen
$form.StartPosition = "CenterScreen"


# ------------------------------------------------------------
# IP ADDRESS LABEL
# ------------------------------------------------------------

# Creates the "Enter IP Address:" label
$label = New-Object System.Windows.Forms.Label

# Text displayed above the search box
$label.Text = "Enter IP Address:"

# Automatically sizes the label to fit the text
$label.AutoSize = $true

# Sets the label position in the window
$label.Location = New-Object System.Drawing.Point(20,25)


# ------------------------------------------------------------
# IP ADDRESS SEARCH BOX
# ------------------------------------------------------------

# Creates the text box where the IP address is entered
$textBox = New-Object System.Windows.Forms.TextBox

# Sets the size of the text box
$textBox.Size = New-Object System.Drawing.Size(250,25)

# Sets the position of the text box
$textBox.Location = New-Object System.Drawing.Point(20,50)


# ------------------------------------------------------------
# BACKSPACE = CLEAR ENTIRE SEARCH BOX
# ------------------------------------------------------------

# Detects when a key is pressed while typing in the IP box
$textBox.Add_KeyDown({

    # Checks whether the Backspace key was pressed
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Back) {

        # Clears the entire IP address instead of deleting one character
        $textBox.Clear()

        # Stops Windows from performing the normal Backspace action
        $_.SuppressKeyPress = $true
    }
})


# ------------------------------------------------------------
# RUN LOOKUP BUTTON
# ------------------------------------------------------------

# Creates the Run Lookup button
$button = New-Object System.Windows.Forms.Button

# Text shown on the button
$button.Text = "Run Lookup"

# Sets the button size
$button.Size = New-Object System.Drawing.Size(100,30)

# Sets the button position
$button.Location = New-Object System.Drawing.Point(285,48)


# ------------------------------------------------------------
# ENTER KEY = RUN LOOKUP
# ------------------------------------------------------------

# Makes pressing Enter perform the same action as clicking Run Lookup
$form.AcceptButton = $button


# ------------------------------------------------------------
# PRIVATE IP ADDRESS REFERENCE
# ------------------------------------------------------------

# Creates the label showing the three private IPv4 ranges
$privateLabel = New-Object System.Windows.Forms.Label

# Text containing the private IPv4 ranges
$privateLabel.Text = "Private IPv4 ranges:`n10.0.0.0 - 10.255.255.255`n172.16.0.0 - 172.31.255.255`n192.168.0.0 - 192.168.255.255"

# Automatically sizes the label
$privateLabel.AutoSize = $true

# Sets the position of the private IP reference
$privateLabel.Location = New-Object System.Drawing.Point(20,90)

# ------------------------------------------------------------
# WEBSITE CHECKLIST
# ------------------------------------------------------------

# Label displayed above the website checklist
$siteLabel = New-Object System.Windows.Forms.Label
$siteLabel.Text = "Lookup Websites:"
$siteLabel.AutoSize = $true
$siteLabel.Location = New-Object System.Drawing.Point(20,160)


# Creates a checklist containing all lookup websites
$siteChecklist = New-Object System.Windows.Forms.CheckedListBox
$siteChecklist.Size = New-Object System.Drawing.Size(365,155)
$siteChecklist.Location = New-Object System.Drawing.Point(20,180)

# One click checks or unchecks a site
$siteChecklist.CheckOnClick = $true

# Removes the blue highlight after clicking a checklist item
$siteChecklist.Add_MouseUp({
    $siteChecklist.ClearSelected()
})


# ------------------------------------------------------------
# LOAD SAVED WEBSITE SETTINGS
# ------------------------------------------------------------

# If settings have previously been saved, load them
$savedSites = $null

if (Test-Path $settingsFile) {
    try {
        $savedSites = Get-Content $settingsFile -Raw | ConvertFrom-Json
    }
    catch {
        $savedSites = $null
    }
}


# Add every website to the checklist
foreach ($site in $sites) {

    # Default to checked
    $shouldBeChecked = $true

    # If saved settings exist, use the previous setting
    if ($null -ne $savedSites) {

        if ($savedSites.PSObject.Properties.Name -contains $site.Name) {
            $shouldBeChecked = [bool]$savedSites.($site.Name)
        }
    }

    # Add the website and set its checked state
    [void]$siteChecklist.Items.Add(
        $site.Name,
        $shouldBeChecked
    )
}

# ------------------------------------------------------------
# LOOKUP HISTORY LABEL
# ------------------------------------------------------------

# Creates the "Lookup History:" label
$historyLabel = New-Object System.Windows.Forms.Label

# Text displayed above the history box
$historyLabel.Text = "Lookup History:"

# Automatically sizes the label
$historyLabel.AutoSize = $true

# Sets the position of the history label
$historyLabel.Location = New-Object System.Drawing.Point(20,350)


# ------------------------------------------------------------
# LOOKUP HISTORY BOX
# ------------------------------------------------------------

# Creates the box that displays previous IP searches
$historyBox = New-Object System.Windows.Forms.ListBox

# Sets the history box size
$historyBox.Size = New-Object System.Drawing.Size(365,140)

# Sets the history box position
$historyBox.Location = New-Object System.Drawing.Point(20,370)


# ------------------------------------------------------------
# CLICK HISTORY ENTRY = PUT IP BACK IN SEARCH BOX
# ------------------------------------------------------------

# Runs whenever a history item is selected
$historyBox.Add_SelectedIndexChanged({

    # Makes sure something was actually selected
    if ($historyBox.SelectedItem) {

        # Gets the complete selected history line
        $selectedEntry = $historyBox.SelectedItem.ToString()

        # Extracts the IP address from the end of the history entry
$historyIP = ($selectedEntry -split '\s+')[-1]

# Checks whether the extracted value is a valid IPv4 or IPv6 address
$parsedIP = $null

if ([System.Net.IPAddress]::TryParse($historyIP, [ref]$parsedIP)) {

    # Places the IP address into the search box
    $textBox.Text = $historyIP

    # Moves keyboard focus back to the search box
    $textBox.Focus()

    # Places the cursor at the end of the IP address
    $textBox.SelectionStart = $textBox.Text.Length
}
    }
})


# ------------------------------------------------------------
# LOAD SAVED HISTORY
# ------------------------------------------------------------

# Checks whether the history file already exists
if (Test-Path $historyFile) {

    # Reads all previous searches from the history file
    $savedHistory = Get-Content $historyFile

    # Adds each saved entry to the visible history box
    foreach ($entry in $savedHistory) {
        $historyBox.Items.Add($entry)
    }
}

# ------------------------------------------------------------
# STYLING
# ------------------------------------------------------------

# Main window background
$form.BackColor = [System.Drawing.Color]::FromArgb(245,248,252)

# Main text color
$label.ForeColor = [System.Drawing.Color]::FromArgb(35,45,55)

# Search box colors
$textBox.BackColor = [System.Drawing.Color]::White
$textBox.ForeColor = [System.Drawing.Color]::FromArgb(25,25,25)
$textBox.BorderStyle = "FixedSingle"

# ------------------------------------------------------------
# RUN LOOKUP BUTTON
# ------------------------------------------------------------

# Makes the button flat instead of the default Windows style
$button.FlatStyle = "Flat"

# Main button color
$button.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)

# Button text color
$button.ForeColor = [System.Drawing.Color]::White

# Removes the default border
$button.FlatAppearance.BorderSize = 0

# Color when hovering over the button
$button.FlatAppearance.MouseOverBackColor = `
    [System.Drawing.Color]::FromArgb(0,100,190)

# Color while clicking the button
$button.FlatAppearance.MouseDownBackColor = `
    [System.Drawing.Color]::FromArgb(0,80,160)



# ------------------------------------------------------------
# PRIVATE IP REFERENCE
# ------------------------------------------------------------

# Gives the private IP section a warm orange color
$privateLabel.ForeColor = [System.Drawing.Color]::FromArgb(190,90,35)

# ------------------------------------------------------------
# LOOKUP WEBSITE SECTION
# ------------------------------------------------------------

# Blue section heading
$siteLabel.ForeColor = [System.Drawing.Color]::FromArgb(35,95,160)

# Checklist colors
$siteChecklist.BackColor = [System.Drawing.Color]::White
$siteChecklist.ForeColor = [System.Drawing.Color]::FromArgb(35,35,35)
$siteChecklist.BorderStyle = "FixedSingle"

# ------------------------------------------------------------
# HISTORY SECTION
# ------------------------------------------------------------

# Blue section heading
$historyLabel.ForeColor = [System.Drawing.Color]::FromArgb(35,95,160)

# History box colors
$historyBox.BackColor = [System.Drawing.Color]::White
$historyBox.ForeColor = [System.Drawing.Color]::FromArgb(35,35,35)
$historyBox.BorderStyle = "FixedSingle"

# ------------------------------------------------------------
# RUN LOOKUP ACTION
# ------------------------------------------------------------

# Everything inside this block happens when Run Lookup is clicked
# or when Enter is pressed
$button.Add_Click({

    # Gets the IP address from the search box
    # Gets whatever the user entered
$inputText = $textBox.Text.Trim()


# --------------------------------------------------------
# NORMALIZE IP INPUT
# --------------------------------------------------------

# Removes http:// or https:// if present
$inputText = $inputText -replace '^https?://', ''

# Removes anything after a forward slash
# Example: 8.8.8.8/test becomes 8.8.8.8
$inputText = $inputText.Split('/')[0]

# Removes a port number from IPv4 addresses
# Example: 8.8.8.8:443 becomes 8.8.8.8
if ($inputText -match '^(\d{1,3}(?:\.\d{1,3}){3}):\d+$') {
    $inputText = $matches[1]
}

# The cleaned result becomes the IP used by the rest of the program
$ip = $inputText.Trim()

# Updates the search box so you can see what was extracted
$textBox.Text = $ip

    # Creates a variable PowerShell will use to validate the IP
    $validIP = $null


    # --------------------------------------------------------
    # VALIDATE IP ADDRESS
    # --------------------------------------------------------

    # Checks whether the entered text is a valid IPv4 or IPv6 address
    if (-not [System.Net.IPAddress]::TryParse($ip, [ref]$validIP)) {

        # Shows an error popup if the IP is invalid
        [System.Windows.Forms.MessageBox]::Show(
            "'$ip' is not a valid IP address.",
            "Invalid IP"
        )

        # Stops the lookup from continuing
        return
    }


    # --------------------------------------------------------
    # CREATE HISTORY ENTRY
    # --------------------------------------------------------

    # Gets the current date and time
    $timestamp = Get-Date -Format "MM/dd/yyyy hh:mm:ss tt"

    # Combines the date/time and IP into one history line
    $historyEntry = "$timestamp   $ip"

    # Adds the newest lookup to the top of the history box
    $historyBox.Items.Insert(0, $historyEntry)


    # --------------------------------------------------------
    # SAVE HISTORY TO FILE
    # --------------------------------------------------------

    # Creates an empty list to hold previous history
    $existingHistory = @()

    # If the history file already exists, read it
    if (Test-Path $historyFile) {
        $existingHistory = Get-Content $historyFile
    }

    # Writes the newest entry first, followed by all older entries
    @($historyEntry) + $existingHistory | Set-Content $historyFile

    # --------------------------------------------------------
    # OPEN SELECTED LOOKUP WEBSITES IN CHROME
    # --------------------------------------------------------

    # Holds all selected lookup URLs
    $selectedUrls = @()

    # Go through every website
    foreach ($site in $sites) {

        # Only use websites that are checked
        if ($siteChecklist.CheckedItems -contains $site.Name) {

            # Insert the IP address into the site's URL
            $url = [string]::Format($site.Url, $ip)

            # Add the finished URL to the list
            $selectedUrls += $url
        }
    }

    # Only continue if at least one website is selected
    if ($selectedUrls.Count -gt 0) {

        # Send all selected URLs to Chrome
        # If Chrome is already open, these should open as new tabs
        Start-Process `
            -FilePath $chromePath `
            -ArgumentList $selectedUrls
    }
    


    # --------------------------------------------------------
    # RESET SEARCH BOX AFTER LOOKUP
    # --------------------------------------------------------

    # Clears the IP that was just searched
    $textBox.Clear()

    # Places the cursor back in the search box
    $textBox.Focus()
})


# ------------------------------------------------------------
# ADD CONTROLS TO WINDOW
# ------------------------------------------------------------

# Adds the IP label to the window
$form.Controls.Add($label)

# Adds the IP search box
$form.Controls.Add($textBox)

# Adds the Run Lookup button
$form.Controls.Add($button)

# Adds the Lookup History label
$form.Controls.Add($historyLabel)

# Adds the Lookup History box
$form.Controls.Add($historyBox)

# Adds the Private IPv4 ranges reference
$form.Controls.Add($privateLabel)

# Adds the Lookup Websites label
$form.Controls.Add($siteLabel)

# Adds the website checklist
$form.Controls.Add($siteChecklist)


# ------------------------------------------------------------
# AUTO-FOCUS SEARCH BOX WHEN WINDOW OPENS
# ------------------------------------------------------------

# Automatically places the cursor in the IP search box
$form.Add_Shown({
    $textBox.Focus()
})

# ------------------------------------------------------------
# SAVE WEBSITE SETTINGS WHEN WINDOW CLOSES
# ------------------------------------------------------------

$form.Add_FormClosing({

    # Creates an object that will store each website's
    # checked or unchecked state
    $settings = [ordered]@{}

    foreach ($site in $sites) {

        # True = checked
        # False = unchecked
        $settings[$site.Name] = (
            $siteChecklist.CheckedItems -contains $site.Name
        )
    }

    # Saves the settings as JSON
    $settings |
        ConvertTo-Json |
        Set-Content $settingsFile
})

# ------------------------------------------------------------
# DISPLAY WINDOW
# ------------------------------------------------------------

# Opens the IP Lookup window and keeps it open until it is closed
[void]$form.ShowDialog()

