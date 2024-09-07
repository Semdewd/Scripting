# Import required modules
Import-Module ActiveDirectory
Import-Module DhcpServer
Import-Module dnsServer
Import-Module IpamServer

# Ensure the script is run with the highest available privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
    Break
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

# Global variables
$script:isLoggedIn = $false
$script:loginTime = Get-Date

# Function to verify AD credentials securely
function Test-ADCredentials {
    param(
        [string]$Username,
        [System.Security.SecureString]$SecurePassword
    )

    $domain = "ds.ad.ssmhc.com"
    $contextType = [System.DirectoryServices.AccountManagement.ContextType]::Domain
    $principalContext = New-Object System.DirectoryServices.AccountManagement.PrincipalContext($contextType, $domain)

    try {
        $credential = New-Object System.Management.Automation.PSCredential($Username, $SecurePassword)
        return $principalContext.ValidateCredentials($credential.UserName, $credential.GetNetworkCredential().Password)
    }
    catch {
        Write-Error "Error validating credentials: $_"
        return $false
    }
    finally {
        if ($principalContext) { $principalContext.Dispose() }
    }
}

# Function to show login dialog
function Show-LoginDialog {
    $loginForm = New-Object System.Windows.Forms.Form
    $loginForm.Text = "Admin Login"
    $loginForm.Size = New-Object System.Drawing.Size(300, 200)
    $loginForm.StartPosition = "CenterScreen"
    $loginForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $loginForm.MaximizeBox = $false
    $loginForm.MinimizeBox = $false

    $usernameLabel = New-Object System.Windows.Forms.Label
    $usernameLabel.Text = "Username:"
    $usernameLabel.Location = New-Object System.Drawing.Point(10, 20)
    $usernameLabel.Size = New-Object System.Drawing.Size(280, 20)
    $loginForm.Controls.Add($usernameLabel)

    $usernameTextBox = New-Object System.Windows.Forms.TextBox
    $usernameTextBox.Location = New-Object System.Drawing.Point(10, 40)
    $usernameTextBox.Size = New-Object System.Drawing.Size(260, 20)
    $loginForm.Controls.Add($usernameTextBox)

    $passwordLabel = New-Object System.Windows.Forms.Label
    $passwordLabel.Text = "Password:"
    $passwordLabel.Location = New-Object System.Drawing.Point(10, 70)
    $passwordLabel.Size = New-Object System.Drawing.Size(280, 20)
    $loginForm.Controls.Add($passwordLabel)

    $passwordTextBox = New-Object System.Windows.Forms.MaskedTextBox
    $passwordTextBox.PasswordChar = '*'
    $passwordTextBox.Location = New-Object System.Drawing.Point(10, 90)
    $passwordTextBox.Size = New-Object System.Drawing.Size(260, 20)
    $loginForm.Controls.Add($passwordTextBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75, 120)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $loginForm.AcceptButton = $okButton
    $loginForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(150, 120)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $loginForm.CancelButton = $cancelButton
    $loginForm.Controls.Add($cancelButton)

    $result = $loginForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $username = $usernameTextBox.Text.Trim()
        if ($username -notmatch '1$') {
            [System.Windows.Forms.MessageBox]::Show("Please login with your admin account (username should end with '1').", "Admin Login Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return @{ LoggedIn = $false }
        }
        $securePassword = ConvertTo-SecureString $passwordTextBox.Text -AsPlainText -Force
        return @{
            Username = $username
            SecurePassword = $securePassword
            LoggedIn = $true
        }
    }
    return @{ LoggedIn = $false }
}

# Function to clone AD groups
function Clone-ADGroups {
    param($sourceObject, $targetObject)
    # ... (rest of the function remains unchanged)
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Admin Tool Application"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"

# Create top pane
$topPane = New-Object System.Windows.Forms.Panel
$topPane.Dock = [System.Windows.Forms.DockStyle]::Top
$topPane.Height = 50

# Admin login button
$adminLoginButton = New-Object System.Windows.Forms.Button
$adminLoginButton.Text = "Admin Login"
$adminLoginButton.Location = New-Object System.Drawing.Point(10, 10)
$adminLoginButton.Add_Click({
    $credentials = Show-LoginDialog
    if ($credentials.LoggedIn) {
        if ($credentials.Username -notmatch '1$') {
            [System.Windows.Forms.MessageBox]::Show("Please login with your admin account (username should end with '1').", "Admin Login Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        elseif (Test-ADCredentials -Username $credentials.Username -SecurePassword $credentials.SecurePassword) {
            [System.Windows.Forms.MessageBox]::Show("Login successful!", "Admin Login")
            $script:loginTime = Get-Date
            $script:isLoggedIn = $true
            $leftPane.Enabled = $true
            $logoutButton.Enabled = $true
            $adminLoginButton.Enabled = $false
        } else {
            [System.Windows.Forms.MessageBox]::Show("Invalid username or password. Please try again.", "Login Failed")
        }
    }
})

# Logout button
$logoutButton = New-Object System.Windows.Forms.Button
$logoutButton.Text = "Logout"
$logoutButton.Location = New-Object System.Drawing.Point(100, 10)
$logoutButton.Enabled = $false
$logoutButton.Add_Click({
    $script:isLoggedIn = $false
    $leftPane.Enabled = $false
    $logoutButton.Enabled = $false
    $adminLoginButton.Enabled = $true
    [System.Windows.Forms.MessageBox]::Show("You have been logged out.", "Logout")
})

# Login timer label
$loginTimerLabel = New-Object System.Windows.Forms.Label
$loginTimerLabel.Text = "Time until next login: 04:00:00"
$loginTimerLabel.Location = New-Object System.Drawing.Point(200, 15)
$loginTimerLabel.AutoSize = $true

$topPane.Controls.Add($adminLoginButton)
$topPane.Controls.Add($logoutButton)
$topPane.Controls.Add($loginTimerLabel)

# Create left pane (TreeView for navigation)
$leftPane = New-Object System.Windows.Forms.TreeView
$leftPane.Dock = [System.Windows.Forms.DockStyle]::Left
$leftPane.Width = 250
$leftPane.BackColor = [System.Drawing.Color]::RoyalBlue
$leftPane.ForeColor = [System.Drawing.Color]::White
$leftPane.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Regular)
$leftPane.Enabled = $false

# Add nodes to TreeView
$adTools = $leftPane.Nodes.Add("AD Tools")
$adTools.Nodes.Add("Clone Workstation")
$adTools.Nodes.Add("Create/Edit Auto Login")
$adTools.Nodes.Add("DHCP Reservation")
$adTools.Nodes.Add("Membership Bundles")
$adTools.Nodes.Add("Next free IP")

$printerScripts = $leftPane.Nodes.Add("Printer Scripts")
$printerScripts.Nodes.Add("Open GPO Packages")
$printerScripts.Nodes.Add("Open TC Printer Folder")
$printerScripts.Nodes.Add("Run Printer Script")

$userScripts = $leftPane.Nodes.Add("User Scripts")
$userScripts.Nodes.Add("Create Generic User")
$userScripts.Nodes.Add("Move User")

$miscellaneous = $leftPane.Nodes.Add("Miscellaneous")
$miscellaneous.Nodes.Add("Create desktop icon")

# Create center pane (main content)
$centerPane = New-Object System.Windows.Forms.Panel
$centerPane.Dock = [System.Windows.Forms.DockStyle]::Fill

# Create right pane (hints and instructions)
$rightPane = New-Object System.Windows.Forms.Panel
$rightPane.Dock = [System.Windows.Forms.DockStyle]::Right
$rightPane.Width = 250
$rightPane.BackColor = [System.Drawing.Color]::LightGray

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Location = New-Object System.Drawing.Point(10, 10)
$hintLabel.Size = New-Object System.Drawing.Size(230, 500)
$hintLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Regular)
$hintLabel.Text = "Select an option from the left pane to see instructions and hints."
$rightPane.Controls.Add($hintLabel)

# Welcome message in center pane
$welcomeLabel = New-Object System.Windows.Forms.Label
$welcomeLabel.Text = "Welcome to the Admin Tool`nSelect an option from the left pane to get started."
$welcomeLabel.Location = New-Object System.Drawing.Point(20, 20)
$welcomeLabel.AutoSize = $true
$centerPane.Controls.Add($welcomeLabel)

# Add controls to the form
$form.Controls.Add($topPane)
$form.Controls.Add($leftPane)
$form.Controls.Add($rightPane)
$form.Controls.Add($centerPane)

# Timer for login countdown
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000 # 1 second
$timer.Add_Tick({
    $now = Get-Date
    $timeDiff = New-TimeSpan -Start $script:loginTime -End $now
    $timeLeft = [TimeSpan]::FromHours(4) - $timeDiff
    
    if ($timeLeft.TotalSeconds -le 0) {
        $loginTimerLabel.Text = "Login required"
        $script:isLoggedIn = $false
        $leftPane.Enabled = $false
        $logoutButton.Enabled = $false
        $adminLoginButton.Enabled = $true
        [System.Windows.Forms.MessageBox]::Show("Your session has expired. Please log in again.", "Session Expired")
    } else {
        $loginTimerLabel.Text = "Time until next login: {0:hh\:mm\:ss}" -f $timeLeft
    }
})
$timer.Start()

# Handle TreeView node clicks
$leftPane.Add_AfterSelect({
    $selectedNode = $leftPane.SelectedNode
    $centerPane.Controls.Clear()
    
    switch ($selectedNode.Text) {

# Inside the TreeView node click event handler, add this case:

"DHCP Reservation" {
    $hintLabel.Text = "Instructions for DHCP Reservation:`n`n1. Select a DHCP scope from the dropdown.`n2. Choose one of the available free IPs.`n3. Enter the device name, MAC address, and description.`n4. Click 'Create Reservation' to add the DHCP reservation.`n`nHint: Ensure the MAC address is in the format AA-BB-CC-DD-EE-FF."

    $centerPane.Controls.Clear()

    # Create a panel to hold the content
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $centerPane.Controls.Add($contentPanel)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "DHCP Reservation"
    $label.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $contentPanel.Controls.Add($label)

    # DHCP Scope dropdown
    $scopeLabel = New-Object System.Windows.Forms.Label
    $scopeLabel.Text = "Select DHCP Scope:"
    $scopeLabel.Location = New-Object System.Drawing.Point(20, 60)
    $scopeLabel.AutoSize = $true
    $contentPanel.Controls.Add($scopeLabel)

    $scopeDropdown = New-Object System.Windows.Forms.ComboBox
    $scopeDropdown.Location = New-Object System.Drawing.Point(60, 100)
    $scopeDropdown.Width = 200
    $contentPanel.Controls.Add($scopeDropdown)

    # Populate DHCP scopes
    $dhcpServer = "S024-DHCP3"  # Replace with your DHCP server name
    $scopes = Get-DhcpServerv4Scope -ComputerName $dhcpServer
    foreach ($scope in $scopes) {
        $scopeDropdown.Items.Add("$($scope.ScopeId) ($($scope.Name))")
    }

    # Free IPs ListBox
    $freeIPsLabel = New-Object System.Windows.Forms.Label
    $freeIPsLabel.Text = "Available Free IPs:"
    $freeIPsLabel.Location = New-Object System.Drawing.Point(20, 110)
    $freeIPsLabel.AutoSize = $true
    $contentPanel.Controls.Add($freeIPsLabel)

    $freeIPsListBox = New-Object System.Windows.Forms.ListBox
    $freeIPsListBox.Location = New-Object System.Drawing.Point(20, 130)
    $freeIPsListBox.Size = New-Object System.Drawing.Size(200, 150)
    $contentPanel.Controls.Add($freeIPsListBox)

    # Device Name
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = "Device Name:"
    $nameLabel.Location = New-Object System.Drawing.Point(250, 60)
    $nameLabel.AutoSize = $true
    $contentPanel.Controls.Add($nameLabel)

    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.Location = New-Object System.Drawing.Point(250, 80)
    $nameTextBox.Width = 200
    $contentPanel.Controls.Add($nameTextBox)

    # MAC Address
    $macLabel = New-Object System.Windows.Forms.Label
    $macLabel.Text = "MAC Address:"
    $macLabel.Location = New-Object System.Drawing.Point(250, 110)
    $macLabel.AutoSize = $true
    $contentPanel.Controls.Add($macLabel)

    $macTextBox = New-Object System.Windows.Forms.TextBox
    $macTextBox.Location = New-Object System.Drawing.Point(250, 130)
    $macTextBox.Width = 200
    $contentPanel.Controls.Add($macTextBox)

    # Description
    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Text = "Description:"
    $descriptionLabel.Location = New-Object System.Drawing.Point(250, 160)
    $descriptionLabel.AutoSize = $true
    $contentPanel.Controls.Add($descriptionLabel)

    $descriptionTextBox = New-Object System.Windows.Forms.TextBox
    $descriptionTextBox.Location = New-Object System.Drawing.Point(250, 180)
    $descriptionTextBox.Width = 200
    $descriptionTextBox.Height = 60
    $descriptionTextBox.Multiline = $true
    $contentPanel.Controls.Add($descriptionTextBox)

    # Create Reservation Button
    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Text = "Create Reservation"
    $createButton.Location = New-Object System.Drawing.Point(250, 250)
    $createButton.Width = 150
    $contentPanel.Controls.Add($createButton)

    # Event handler for scope selection
    $scopeDropdown.Add_SelectedIndexChanged({
        $selectedScope = $scopeDropdown.SelectedItem.ToString().Split()[0]
        $freeIPsListBox.Items.Clear()
        
        $scope = Get-DhcpServerv4Scope -ComputerName $dhcpServer -ScopeId $selectedScope
        $startIP = [System.Net.IPAddress]::Parse($scope.StartRange)
        $endIP = [System.Net.IPAddress]::Parse($scope.EndRange)
        $currentIP = $startIP
        $freeIPCount = 0

        while ($currentIP.Address -le $endIP.Address -and $freeIPCount -lt 10) {
            $ip = $currentIP.ToString()
            if (Test-FreeIP -IP $ip -DHCPServer $dhcpServer -DNSServer "S024-DC1" -DHCPScopes $scopes) {
                $freeIPsListBox.Items.Add($ip)
                $freeIPCount++
            }
            $currentIP = [System.Net.IPAddress]::Parse($currentIP.Address + 1)
        }
    })

    # Event handler for Create Reservation button
    $createButton.Add_Click({
        $selectedScope = $scopeDropdown.SelectedItem.ToString().Split()[0]
        $selectedIP = $freeIPsListBox.SelectedItem
        $deviceName = $nameTextBox.Text.Trim()
        $macAddress = $macTextBox.Text.Trim()
        $description = $descriptionTextBox.Text.Trim()

        if ([string]::IsNullOrEmpty($selectedScope) -or [string]::IsNullOrEmpty($selectedIP) -or 
            [string]::IsNullOrEmpty($deviceName) -or [string]::IsNullOrEmpty($macAddress)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all required fields.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        try {
            Add-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $selectedScope -IPAddress $selectedIP -ClientId $macAddress -Name $deviceName -Description $description
            [System.Windows.Forms.MessageBox]::Show("DHCP Reservation created successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Clear fields and refresh free IPs
            $nameTextBox.Clear()
            $macTextBox.Clear()
            $descriptionTextBox.Clear()
            $scopeDropdown.SelectedIndex = -1
            $freeIPsListBox.Items.Clear()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating DHCP Reservation: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
}

		
"Membership Bundles" {
    $hintLabel.Text = "Instructions for Membership Bundles:`n`n1. Enter the target object name (computer or user).`n2. Click on a bundle button to apply the corresponding membership bundle to the specified object.`n3. Use the arrows to move items between the list boxes.`n4. Click 'Submit' to add the selected groups to the target object.`n`nHint: Ensure you have the necessary permissions to modify AD group memberships."

    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $centerPane.Controls.Add($contentPanel)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Membership Bundles"
    $label.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(350, 80)

    # Add a "Target Object" label and text box
$targetObjectLabel = New-Object System.Windows.Forms.Label
$targetObjectLabel.Text = "Target Object:"
$targetObjectLabel.AutoSize = $true
$targetObjectLabel.Location = New-Object System.Drawing.Point(250, ($label.Bottom + 20))
$contentPanel.Controls.Add($targetObjectLabel)

$script:BundleTarget = New-Object System.Windows.Forms.TextBox
$script:BundleTarget.Width = 200
$script:BundleTarget.Height = 25
$script:BundleTarget.Font = New-Object System.Drawing.Font("Arial", 10)
$script:BundleTarget.Location = New-Object System.Drawing.Point(($targetObjectLabel.Right + 10), $targetObjectLabel.Top)
$contentPanel.Controls.Add($script:BundleTarget)


    $BundleTarget = New-Object System.Windows.Forms.TextBox
    $BundleTarget.Width = 200
    $BundleTarget.Height = 25
    $BundleTarget.Font = New-Object System.Drawing.Font("Arial", 10)
    $BundleTarget.Location = New-Object System.Drawing.Point(($targetObjectLabel.Right + 10), $targetObjectLabel.Top)
    $contentPanel.Controls.Add($BundleTarget)

    $label.Left = ($contentPanel.ClientSize.Width - $label.Width) / 2
    $label.Top = 80

    $buttonNames = @(
        "Clinic Exam Room",
        "Hospital Exam Room",
        "Knowledge Worker",
        "Nurse Station Type 1",
		"Nurse Station Type 2",
        "Physician Station",
        "Registration Station",
		"Remote Station"
    )

    $buttonWidth = 135
    $buttonHeight = 30
    $buttonsPerRow = 3
    $horizontalSpacing = 20
    $verticalSpacing = 20

    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.AutoSize = $true
    $contentPanel.Controls.Add($buttonPanel)

    $script:bundleContents = @{
        "Clinic Exam Room" = @("G024-EnableBitlocker", "G024-GPO-OneSign - Computer","G024-SCCM-Workstation-Servicing-Ring3a - Computer","G024-SWD-ImprivataAgent_x64_1User - Computer")
        "Hospital Exam Room" = @("G024-Citrix-Epic", "G024-Onesign-migration", "G024-SWD-Instamed")
        "Knowledge Worker" = @("G024-EnableBitlocker","G024-GPO-OneSign - Computer","G024-SCCM-Workstation-Servicing-Ring3d - Computer","G024-SWD-ImprivataAgent_x64_1User - Computer","G024-SWD-O365_ProPlus_x64 - Computer")
        "Nurse Station Type 1" = @("G024-EnableBitlocker", "G024-GPO-OneSign - Computer", "G024-SCCM-Workstation-Servicing-Ring3a - Computer","G024-SWD-ImprivataAgent_x64_1User - Computer","G024-SWD-SkypeForBusinessBasic_2016_x64_VI_TEST - Computer")
		"Nurse Station Type 2" = @("G024-EnableBitlocker", "G024-GPO-OneSign - Computer", "G024-SCCM-Workstation-Servicing-Ring3a - Computer","G024-SWD-ImprivataAgent_x64_Kiosk - Computer","G024-SWD-SkypeForBusinessBasic_2016_x64_VI_TEST - Computer")
        "Physician Station" = @("G024-EnableBitlocker","G024-GPO-OneSign - Computer", "G024-SWD-Fluency","G024-SWD-FluencyDirect_VI - Computer","G024-SWD-FluencyFix - Computer","G024-SWD-ImprivataAgent_x64_1User - Computer","G024-SWD-RA1000_609_TCH_OK_VI - Computer","G024-SWD-O365_ProPlus_x64 - Computer")
        "Registration Station" = @("G024-EnableBitlocker", "G024-GPO-OneSign - Computer", "G024-SCCM-Workstation-Servicing-Ring3a - Computer", "G024-SWD-ImprivataAgent_x64_1User - Computer","G024-SWD-IngenicoUSBDriver - Computer","G024-SWD-InstaMedDOTNETAPI2.8-R01 - Computer","G024-SWD-PaperStream_IP_TWAIN_1.60.0 - Computer","G024-SWD-O365_ProPlus_x64 - Computer","G024-GPO-ScreenSaver Timeout 20 Minutes - User")
        "Remote Station" = @("G024-EnableBitlocker","G024-AZ-AlwaysOnVPN-SSM - Computer", "G024-SCCM-Workstation-Servicing-Ring3a - Computer","G024-SWD-PaperStream_IP_TWAIN_1.60.0 - Computer","G024-SWD-O365_ProPlus_x64 - Computer")
	}

    for ($i = 0; $i -lt $buttonNames.Count; $i++) {
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $buttonNames[$i]
        $button.Width = $buttonWidth
        $button.Height = $buttonHeight
        $row = [Math]::Floor($i / $buttonsPerRow)
        $col = $i % $buttonsPerRow
        $x = ($buttonWidth + $horizontalSpacing) * $col
        $y = ($buttonHeight + $verticalSpacing) * $row
        $button.Location = New-Object System.Drawing.Point($x, $y)
        $button.Add_Click({
            $bundleName = $this.Text
            $bundleGroups = $script:bundleContents[$bundleName]
            
            if ($bundleGroups) {
                $itemsMoved = 0
                foreach ($group in $bundleGroups) {
                    if ($script:listBox1.Items.Contains($group)) {
                        $script:listBox1.Items.Remove($group)
                        $script:listBox2.Items.Add($group)
                        $itemsMoved++
                    }
                }
                
                if ($itemsMoved -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("No AD Groups to move!", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }
            } else {
                [System.Windows.Forms.MessageBox]::Show("Bundle content not found for $bundleName", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        })
        $buttonPanel.Controls.Add($button)
    }

    $buttonPanel.Left = ($contentPanel.ClientSize.Width - $buttonPanel.Width) / 2
    $buttonPanel.Top = 180

    $script:listBox1 = New-Object System.Windows.Forms.ListBox
    $script:listBox1.Width = 200
    $script:listBox1.Height = 200
    $script:listBox1.Location = New-Object System.Drawing.Point(($contentPanel.ClientSize.Width / 2 - 230), ($buttonPanel.Bottom + 20))
    $contentPanel.Controls.Add($script:listBox1)

    $script:listBox1.Items.AddRange(@(
        "G024-AZ-AlwaysOnVPN-SSM - Computer",
		"G024-EnableBitlocker",
        "G024-GPO-OneSign - Computer",
        "G024-SWD-ImprivataAgent_x64_1User - Computer",
		"G024-SWD-IngenicoUSBDriver - Computer",
		"G024-SWD-InstaMedDOTNETAPI2.8-R01 - Computer",
        "G024-SWD-FluencyDirect_VI - Computer",
		"G024-SWD-FluencyFix - Computer",
		"G024-SWD-O365_ProPlus_x64 - Computer",
		"G024-SWD-PaperStream_IP_TWAIN_1.60.0 - Computer",
		"G024-SWD-RA1000_609_TCH_OK_VI - Computer",
		"G024-SWD-SkypeForBusinessBasic_2016_x64_VI_TEST - Computer",
        "G024-SCCM-Workstation-Servicing-Ring3a - Computer",
		"G024-SCCM-Workstation-Servicing-Ring3d - Computer"
    ))

    $script:listBox2 = New-Object System.Windows.Forms.ListBox
    $script:listBox2.Width = 200
    $script:listBox2.Height = 200
    $script:listBox2.Location = New-Object System.Drawing.Point(($contentPanel.ClientSize.Width / 2 + 30), ($buttonPanel.Bottom + 20))
    $contentPanel.Controls.Add($script:listBox2)

    $label1 = New-Object System.Windows.Forms.Label
    $label1.Text = "Available Groups"
    $label1.AutoSize = $true
    $label1.Location = New-Object System.Drawing.Point($script:listBox1.Left, ($script:listBox1.Top - 20))
    $contentPanel.Controls.Add($label1)

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Text = "Selected Groups"
    $label2.AutoSize = $true
    $label2.Location = New-Object System.Drawing.Point($script:listBox2.Left, ($script:listBox2.Top - 20))
    $contentPanel.Controls.Add($label2)

    $moveRightButton = New-Object System.Windows.Forms.Button
    $moveRightButton.Text = ">"
    $moveRightButton.Width = 40
    $moveRightButton.Height = 30
    $moveRightButton.Location = New-Object System.Drawing.Point(($contentPanel.ClientSize.Width / 2 - 20), ($script:listBox1.Top + 50))
    $contentPanel.Controls.Add($moveRightButton)

    $moveLeftButton = New-Object System.Windows.Forms.Button
    $moveLeftButton.Text = "<"
    $moveLeftButton.Width = 40
    $moveLeftButton.Height = 30
    $moveLeftButton.Location = New-Object System.Drawing.Point(($contentPanel.ClientSize.Width / 2 - 20), ($script:listBox1.Top + 90))
    $contentPanel.Controls.Add($moveLeftButton)

    $moveRightButton.Add_Click({
        $selectedItems = @($script:listBox1.SelectedItems)
        foreach ($item in $selectedItems) {
            $script:listBox2.Items.Add($item)
        }
        foreach ($item in $selectedItems) {
            $script:listBox1.Items.Remove($item)
        }
    })

    $moveLeftButton.Add_Click({
        $selectedItems = @($script:listBox2.SelectedItems)
        foreach ($item in $selectedItems) {
            $script:listBox1.Items.Add($item)
        }
        foreach ($item in $selectedItems) {
            $script:listBox2.Items.Remove($item)
        }
    })

    $submitButton = New-Object System.Windows.Forms.Button
    $submitButton.Text = "Submit"
    $submitButton.Width = 100
    $submitButton.Height = 30
    $submitButton.Location = New-Object System.Drawing.Point(
        [int](($contentPanel.ClientSize.Width - $submitButton.Width) / 2),
        ($script:listBox1.Bottom + 10)
    )
    $contentPanel.Controls.Add($submitButton)

   $submitButton.Add_Click({
    $targetObject = $script:BundleTarget.Text.Trim()
    $selectedGroups = $script:listBox2.Items

    if ([string]::IsNullOrEmpty($targetObject)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a target object name.", "Missing Target", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }


        if ([string]::IsNullOrEmpty($targetObject)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a target object name.", "Missing Target", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if ($selectedGroups.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one group before submitting.", "No Groups Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        try {
            $adObject = Get-ADObject -Filter {(Name -eq $targetObject) -or (SamAccountName -eq $targetObject)} -Properties MemberOf
            if (-not $adObject) {
                throw "Target object not found in Active Directory."
            }

            $addedGroups = @()
            $failedGroups = @()

            foreach ($group in $selectedGroups) {
                try {
                    $adGroup = Get-ADGroup -Identity $group
                    Add-ADGroupMember -Identity $adGroup -Members $adObject -ErrorAction Stop
                    $addedGroups += $group
                }
                catch {
                    $failedGroups += $group
                    Write-Host "Failed to add $targetObject to group $group. Error: $_"
                }
            }

            $message = "Groups added successfully: " + ($addedGroups -join ", ")
            if ($failedGroups.Count -gt 0) {
                $message += "`n`nFailed to add groups: " + ($failedGroups -join ", ")
            }

            [System.Windows.Forms.MessageBox]::Show($message, "Group Assignment Result", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error: $_", "Group Assignment Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = "Reset"
    $resetButton.Width = 100
    $resetButton.Height = 30
    $resetButton.Location = New-Object System.Drawing.Point(
        [int](($contentPanel.ClientSize.Width - $resetButton.Width - $submitButton.Width - 10) / 2),
        ($script:listBox1.Bottom + 10)
    )
    $contentPanel.Controls.Add($resetButton)

    $submitButton.Location = New-Object System.Drawing.Point(
        ($resetButton.Right + 10),
        $resetButton.Top
    )

    $resetButton.Add_Click({
        $itemsToMove = $script:listBox2.Items.Clone()
        foreach ($item in $itemsToMove) {
            $script:listBox1.Items.Add($item)
            $script:listBox2.Items.Remove($item)
        }
    })

    $contentPanel.Add_Resize({
        $label.Left = ($contentPanel.Width - $label.Width) / 2
        $targetObjectLabel.Left = ($contentPanel.Width - $targetObjectLabel.Width - $BundleTarget.Width - 10) / 2
        $BundleTarget.Left = $targetObjectLabel.Right + 10
        $buttonPanel.Left = ($contentPanel.Width - $buttonPanel.Width) / 2
        $script:listBox1.Left = ($contentPanel.Width / 2 - 230)
        $script:listBox2.Left = ($contentPanel.Width / 2 + 30)
        $label1.Left = $script:listBox1.Left
        $label2.Left = $script:listBox2.Left
        $moveRightButton.Left = ($contentPanel.Width / 2 - 20)
        $moveLeftButton.Left = ($contentPanel.Width / 2 - 20)
        $submitButton.Left = ($contentPanel.Width - $submitButton.Width) / 2
        $submitButton.Top = $script:listBox1.Bottom + 10
        $resetButton.Left = ($contentPanel.Width - $resetButton.Width - $submitButton.Width - 10) / 2
        $resetButton.Top = $script:listBox1.Bottom + 10
        $submitButton.Left = $resetButton.Right + 10
    })
}


# Inside the TreeView node click event handler, update the "DHCP Reservation" case:

"DHCP Reservation" {
    $hintLabel.Text = "Instructions for DHCP Reservation:`n`n1. Select a DHCP server from the dropdown.`n2. Choose a DHCP scope for the selected server.`n3. Select one of the available free IPs.`n4. Enter the Reservation Name, MAC address, and description.`n5. Click 'Create Reservation' to add the DHCP reservation.`n`nHint: Ensure the MAC address is in the format AA-BB-CC-DD-EE-FF."

    $centerPane.Controls.Clear()

    # Create a panel to hold the content
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $centerPane.Controls.Add($contentPanel)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "DHCP Reservation"
    $label.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $contentPanel.Controls.Add($label)

    # DHCP Server dropdown
    $serverLabel = New-Object System.Windows.Forms.Label
    $serverLabel.Text = "Select DHCP Server:"
    $serverLabel.Location = New-Object System.Drawing.Point(10, 60)
    $serverLabel.AutoSize = $true
    $contentPanel.Controls.Add($serverLabel)

    $serverDropdown = New-Object System.Windows.Forms.ComboBox
    $serverDropdown.Location = New-Object System.Drawing.Point(20, 80)
    $serverDropdown.Width = 200
    $contentPanel.Controls.Add($serverDropdown)

    # Populate DHCP servers
    $dhcpServers = Get-DhcpServerInDC | Select-Object -ExpandProperty DnsName
    foreach ($server in $dhcpServers) {
        $serverDropdown.Items.Add($server)
    }

    # DHCP Scope dropdown
    $scopeLabel = New-Object System.Windows.Forms.Label
    $scopeLabel.Text = "Select DHCP Scope:"
    $scopeLabel.Location = New-Object System.Drawing.Point(20, 110)
    $scopeLabel.AutoSize = $true
    $contentPanel.Controls.Add($scopeLabel)

    $scopeDropdown = New-Object System.Windows.Forms.ComboBox
    $scopeDropdown.Location = New-Object System.Drawing.Point(20, 130)
    $scopeDropdown.Width = 200
    $contentPanel.Controls.Add($scopeDropdown)

    # Free IPs ListBox
    $freeIPsLabel = New-Object System.Windows.Forms.Label
    $freeIPsLabel.Text = "Available Free IPs:"
    $freeIPsLabel.Location = New-Object System.Drawing.Point(20, 160)
    $freeIPsLabel.AutoSize = $true
    $contentPanel.Controls.Add($freeIPsLabel)

    $freeIPsListBox = New-Object System.Windows.Forms.ListBox
    $freeIPsListBox.Location = New-Object System.Drawing.Point(20, 180)
    $freeIPsListBox.Size = New-Object System.Drawing.Size(200, 150)
    $contentPanel.Controls.Add($freeIPsListBox)

    # Reservation Name
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = "Reservation Name:"
    $nameLabel.Location = New-Object System.Drawing.Point(250, 60)
    $nameLabel.AutoSize = $true
    $contentPanel.Controls.Add($nameLabel)

    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.Location = New-Object System.Drawing.Point(250, 80)
    $nameTextBox.Width = 200
    $contentPanel.Controls.Add($nameTextBox)

    # IP Address
    $ipLabel = New-Object System.Windows.Forms.Label
    $ipLabel.Text = "IP Address:"
    $ipLabel.Location = New-Object System.Drawing.Point(250, 110)
    $ipLabel.AutoSize = $true
    $contentPanel.Controls.Add($ipLabel)

    $ipTextBox = New-Object System.Windows.Forms.TextBox
    $ipTextBox.Location = New-Object System.Drawing.Point(250, 130)
    $ipTextBox.Width = 200
    $ipTextBox.ReadOnly = $true
    $contentPanel.Controls.Add($ipTextBox)

    # MAC Address
    $macLabel = New-Object System.Windows.Forms.Label
    $macLabel.Text = "MAC Address:"
    $macLabel.Location = New-Object System.Drawing.Point(250, 160)
    $macLabel.AutoSize = $true
    $contentPanel.Controls.Add($macLabel)

    $macTextBox = New-Object System.Windows.Forms.TextBox
    $macTextBox.Location = New-Object System.Drawing.Point(250, 180)
    $macTextBox.Width = 200
    $contentPanel.Controls.Add($macTextBox)

    # Description
    $descriptionLabel = New-Object System.Windows.Forms.Label
    $descriptionLabel.Text = "Description:"
    $descriptionLabel.Location = New-Object System.Drawing.Point(250, 210)
    $descriptionLabel.AutoSize = $true
    $contentPanel.Controls.Add($descriptionLabel)

    $descriptionTextBox = New-Object System.Windows.Forms.TextBox
    $descriptionTextBox.Location = New-Object System.Drawing.Point(250, 230)
    $descriptionTextBox.Width = 200
    $descriptionTextBox.Height = 60
    $descriptionTextBox.Multiline = $true
    $contentPanel.Controls.Add($descriptionTextBox)

    # Create Reservation Button
    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Text = "Create Reservation"
    $createButton.Location = New-Object System.Drawing.Point(250, 300)
    $createButton.Width = 150
    $contentPanel.Controls.Add($createButton)

    # Event handler for server selection
    $serverDropdown.Add_SelectedIndexChanged({
        $selectedServer = $serverDropdown.SelectedItem
        $scopeDropdown.Items.Clear()
        $scopes = Get-DhcpServerv4Scope -ComputerName $selectedServer
        foreach ($scope in $scopes) {
            $scopeDropdown.Items.Add("$($scope.ScopeId) ($($scope.Name))")
        }
    })

    # Event handler for scope selection
    $scopeDropdown.Add_SelectedIndexChanged({
        $selectedServer = $serverDropdown.SelectedItem
        $selectedScope = $scopeDropdown.SelectedItem.ToString().Split()[0]
        $freeIPsListBox.Items.Clear()
        
        $scope = Get-DhcpServerv4Scope -ComputerName $selectedServer -ScopeId $selectedScope
        $startIP = [System.Net.IPAddress]::Parse($scope.StartRange)
        $endIP = [System.Net.IPAddress]::Parse($scope.EndRange)
        $currentIP = $startIP
        $freeIPCount = 0

        while ($currentIP.Address -le $endIP.Address -and $freeIPCount -lt 10) {
            $ip = $currentIP.ToString()
            if (Test-FreeIP -IP $ip -DHCPServer $selectedServer -DNSServer "YourDNSServer" -DHCPScopes $scope) {
                $freeIPsListBox.Items.Add($ip)
                $freeIPCount++
            }
            $currentIP = [System.Net.IPAddress]::Parse($currentIP.Address + 1)
        }
    })

    # Event handler for free IP selection
    $freeIPsListBox.Add_SelectedIndexChanged({
        $ipTextBox.Text = $freeIPsListBox.SelectedItem
    })

    # Event handler for Create Reservation button
    $createButton.Add_Click({
        $selectedServer = $serverDropdown.SelectedItem
        $selectedScope = $scopeDropdown.SelectedItem.ToString().Split()[0]
        $reservationName = $nameTextBox.Text.Trim()
        $ipAddress = $ipTextBox.Text.Trim()
        $macAddress = $macTextBox.Text.Trim()
        $description = $descriptionTextBox.Text.Trim()

        if ([string]::IsNullOrEmpty($selectedServer) -or [string]::IsNullOrEmpty($selectedScope) -or 
            [string]::IsNullOrEmpty($reservationName) -or [string]::IsNullOrEmpty($ipAddress) -or 
            [string]::IsNullOrEmpty($macAddress)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all required fields.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        try {
            Add-DhcpServerv4Reservation -ComputerName $selectedServer -ScopeId $selectedScope -IPAddress $ipAddress -ClientId $macAddress -Name $reservationName -Description $description
            [System.Windows.Forms.MessageBox]::Show("DHCP Reservation created successfully.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Clear fields and refresh free IPs
            $nameTextBox.Clear()
            $ipTextBox.Clear()
            $macTextBox.Clear()
            $descriptionTextBox.Clear()
            $scopeDropdown.SelectedIndex = -1
            $freeIPsListBox.Items.Clear()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating DHCP Reservation: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
}

		
		"Create/Edit Auto Login" {
            $hintLabel.Text = "Instructions for Create/Edit Auto Login:`n`n1. Click the 'Run Auto Login Tool' button to execute the SetAutoLogon-v2.hta file.`n`nHint: The tool will run in the background. Check for any output or errors in the console."
            
            $centerPane.Controls.Clear()
            
            # Create a panel to hold the content
            $contentPanel = New-Object System.Windows.Forms.Panel
            $contentPanel.Width = 400
            $contentPanel.Height = 200
            $contentPanel.Left = ($centerPane.ClientSize.Width - $contentPanel.Width) / 2
            $contentPanel.Top = ($centerPane.ClientSize.Height - $contentPanel.Height) / 2
            $centerPane.Controls.Add($contentPanel)

            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Create/Edit Auto Login"
            $label.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
            $label.AutoSize = $true
            $label.Location = New-Object System.Drawing.Point(20, 20)

            $runButton = New-Object System.Windows.Forms.Button
            $runButton.Text = "Run Auto Login Tool"
            $runButton.Location = New-Object System.Drawing.Point(20, 60)
            $runButton.Width = 200
            $runButton.Add_Click({
                # Use $PSScriptRoot to get the directory of the current script
                $scriptPath = Join-Path $PSScriptRoot "SetAutoLogon-v2.hta"
                
                if (Test-Path $scriptPath) {
                    try {
                        # Run the HTA file directly
                        $output = & mshta.exe $scriptPath
                        
                        # Display the output in a message box
                        if ($output) {
                            [System.Windows.Forms.MessageBox]::Show("Auto Login Tool executed. Output:`n`n$output", "Tool Executed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        } else {
                            [System.Windows.Forms.MessageBox]::Show("Auto Login Tool executed successfully.", "Tool Executed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        }
                    }
                    catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to run the Auto Login Tool. Error: $($_.Exception.Message)", "Tool Execution Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show("The Auto Login Tool file was not found in the application directory. Expected path: $scriptPath", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            })

            $contentPanel.Controls.AddRange(@($label, $runButton))
        }
		
		"Run Printer Script" {
            $hintLabel.Text = "Instructions for Run Printer Script:`n`n1. Click the 'Run Printer Script' button to execute the LaunchAssignPrintersToWorkstationHTA.vbs script.`n`nHint: The script will run in the background. Check for any output or errors in the console."
            
            $centerPane.Controls.Clear()
            
            # Create a panel to hold the content
            $contentPanel = New-Object System.Windows.Forms.Panel
            $contentPanel.Width = 400
            $contentPanel.Height = 200
            $contentPanel.Left = ($centerPane.ClientSize.Width - $contentPanel.Width) / 2
            $contentPanel.Top = ($centerPane.ClientSize.Height - $contentPanel.Height) / 2
            $centerPane.Controls.Add($contentPanel)

            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Run Printer Script"
            $label.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
            $label.AutoSize = $true
            $label.Location = New-Object System.Drawing.Point(20, 20)

            $runButton = New-Object System.Windows.Forms.Button
            $runButton.Text = "Run Printer Script"
            $runButton.Location = New-Object System.Drawing.Point(20, 60)
            $runButton.Width = 200
            $runButton.Add_Click({
                # Use $PSScriptRoot to get the directory of the current script
                $scriptPath = Join-Path $PSScriptRoot "LaunchAssignPrintersToWorkstationHTA.vbs"
                
                if (Test-Path $scriptPath) {
                    try {
                        # Run the VBS script directly
                        $output = & cscript.exe //NoLogo $scriptPath
                        
                        # Display the output in a message box
                        [System.Windows.Forms.MessageBox]::Show("Printer script executed. Output:`n`n$output", "Script Executed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    }
                    catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to run the printer script. Error: $($_.Exception.Message)", "Script Execution Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    }
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show("The printer script file was not found in the application directory. Expected path: $scriptPath", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            })

            $contentPanel.Controls.AddRange(@($label, $runButton))
        }
		
        "Clone Workstation" {     
            $hintLabel.Text = "Instructions for Clone Workstation:`n`n1. Enter the name of the source PC to clone.`n2. Enter the name of the target PC to be cloned.`n3. Click 'Submit' or press Enter to clone the workstation.`n`nHint: Ensure both PC names are correct and exist in Active Directory."
            
            $centerPane.Controls.Clear()
            # Create a panel to hold the content
            $contentPanel = New-Object System.Windows.Forms.Panel
            $contentPanel.Width = 400
            $contentPanel.Height = 500
            $contentPanel.Left = ($centerPane.ClientSize.Width - $contentPanel.Width) / 2
            $contentPanel.Top = ($centerPane.ClientSize.Height - $contentPanel.Height) / 2
            $centerPane.Controls.Add($contentPanel)

            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Clone Workstation"
            $label.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
            $label.AutoSize = $true
            $label.Location = New-Object System.Drawing.Point(20, 20)

            $sourcePCLabel = New-Object System.Windows.Forms.Label
            $sourcePCLabel.Text = "Source PC to clone:"
            $sourcePCLabel.Location = New-Object System.Drawing.Point(20, 60)
            $sourcePCLabel.AutoSize = $true

            $script:sourcePCTextBox = New-Object System.Windows.Forms.TextBox
            $script:sourcePCTextBox.Location = New-Object System.Drawing.Point(20, 80)
            $script:sourcePCTextBox.Width = 360

            $targetPCLabel = New-Object System.Windows.Forms.Label
            $targetPCLabel.Text = "Target PC to be cloned:"
            $targetPCLabel.Location = New-Object System.Drawing.Point(20, 110)
            $targetPCLabel.AutoSize = $true

            $script:targetPCTextBox = New-Object System.Windows.Forms.TextBox
            $script:targetPCTextBox.Location = New-Object System.Drawing.Point(20, 130)
            $script:targetPCTextBox.Width = 360

            $submitButton = New-Object System.Windows.Forms.Button
            $submitButton.Text = "Submit"
            $submitButton.Width = 100
            $submitButton.Location = New-Object System.Drawing.Point(150, 160)
            
            $cloneWorkstation = {
                $script:sourcePC = $script:sourcePCTextBox.Text
                $script:targetPC = $script:targetPCTextBox.Text
                if ([string]::IsNullOrWhiteSpace($script:sourcePC) -or [string]::IsNullOrWhiteSpace($script:targetPC)) {
                    [System.Windows.Forms.MessageBox]::Show("Please enter both Source PC and Target PC names.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                } else {
                    Clone-Workstation $script:sourcePC $script:targetPC
                }
            }
            $submitButton.Add_Click($cloneWorkstation)

            $script:unableToAddBox = New-Object System.Windows.Forms.ListBox
            $script:unableToAddBox.Location = New-Object System.Drawing.Point(20, 220)
            $script:unableToAddBox.Size = New-Object System.Drawing.Size(170, 200)
            $script:unableToAddBox.ForeColor = [System.Drawing.Color]::Red

            $unableToAddLabel = New-Object System.Windows.Forms.Label
            $unableToAddLabel.Text = "Unable to add these memberships:"
            $unableToAddLabel.Location = New-Object System.Drawing.Point(20, 200)
            $unableToAddLabel.AutoSize = $true

            $script:addedBox = New-Object System.Windows.Forms.ListBox
            $script:addedBox.Location = New-Object System.Drawing.Point(210, 220)
            $script:addedBox.Size = New-Object System.Drawing.Size(170, 200)

            $addedLabel = New-Object System.Windows.Forms.Label
            $addedLabel.Text = "Memberships added:"
            $addedLabel.Location = New-Object System.Drawing.Point(210, 200)
            $addedLabel.AutoSize = $true

            $contentPanel.Controls.AddRange(@($label, $sourcePCLabel, $script:sourcePCTextBox, $targetPCLabel, $script:targetPCTextBox, $submitButton, $unableToAddLabel, $script:unableToAddBox, $addedLabel, $script:addedBox))

            # Handle Enter key press
            $form.AcceptButton = $submitButton

            # Set focus to the source PC textbox
            $script:sourcePCTextBox.Select()
        }

"Next free IP" {
    $hintLabel.Text = "Instructions for Next Free IP:`n`n1. Enter the Start IP address and End IP address.`n2. Click 'Find Free IPs' to search for available IP addresses.`n3. The first 10 free IPs will be listed.`n4. If no free IPs are found, IPs with DNS records but no DHCP reservation will be shown.`n`nHint: Enter IP addresses in the format xxx.xxx.xxx.xxx"

    $centerPane.Controls.Clear()

    # Create a panel to hold the content
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Width = 400
    $contentPanel.Height = 500
    $contentPanel.Left = ($centerPane.ClientSize.Width - $contentPanel.Width) / 2
    $contentPanel.Top = ($centerPane.ClientSize.Height - $contentPanel.Height) / 2
    $centerPane.Controls.Add($contentPanel)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Find Next Free IP"
    $label.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $contentPanel.Controls.Add($label)

    $startIPLabel = New-Object System.Windows.Forms.Label
    $startIPLabel.Text = "Start IP Address:"
    $startIPLabel.Location = New-Object System.Drawing.Point(20, 60)
    $startIPLabel.AutoSize = $true
    $contentPanel.Controls.Add($startIPLabel)

    $script:startIPTextBox = New-Object System.Windows.Forms.TextBox
    $script:startIPTextBox.Location = New-Object System.Drawing.Point(20, 85)
    $script:startIPTextBox.Width = 150
    $contentPanel.Controls.Add($script:startIPTextBox)

    $endIPLabel = New-Object System.Windows.Forms.Label
    $endIPLabel.Text = "End IP Address:"
    $endIPLabel.Location = New-Object System.Drawing.Point(200, 60)
    $endIPLabel.AutoSize = $true
    $contentPanel.Controls.Add($endIPLabel)

    $script:endIPTextBox = New-Object System.Windows.Forms.TextBox
    $script:endIPTextBox.Location = New-Object System.Drawing.Point(200, 85)
    $script:endIPTextBox.Width = 150
    $contentPanel.Controls.Add($script:endIPTextBox)

    $findButton = New-Object System.Windows.Forms.Button
    $findButton.Text = "Find Free IPs"
    $findButton.Location = New-Object System.Drawing.Point(20, 115)
    $findButton.Width = 120
    $contentPanel.Controls.Add($findButton)

    $script:resultListBox = New-Object System.Windows.Forms.ListBox
    $script:resultListBox.Location = New-Object System.Drawing.Point(20, 150)
    $script:resultListBox.Size = New-Object System.Drawing.Size(360, 330)
    $contentPanel.Controls.Add($script:resultListBox)

    $findButton.Add_Click({
        if ($null -eq $script:startIPTextBox -or $null -eq $script:endIPTextBox) {
            [System.Windows.Forms.MessageBox]::Show("TextBox controls are not initialized properly.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        $startIP = $script:startIPTextBox.Text.Trim()
        $endIP = $script:endIPTextBox.Text.Trim()

        # IP address validation function
        function Is-ValidIP($ip) {
            return $ip -match "^\d{1,3}(\.\d{1,3}){3}$" -and ($ip.Split(".") | ForEach-Object { $_ -ge 0 -and $_ -le 255 })
        }

        $invalidIPs = @()
        if (-not (Is-ValidIP $startIP)) { $invalidIPs += "Start IP" }
        if (-not (Is-ValidIP $endIP)) { $invalidIPs += "End IP" }

        if ($invalidIPs.Count -eq 0) {
            $script:resultListBox.Items.Clear()
            $script:resultListBox.Items.Add("Searching for free IPs... This may take a while.")
            $script:resultListBox.Refresh()

            # Get DHCP server
            $dhcpServer = "S024-DHCP3"  # Replace with your DHCP server name or IP
            
            # Ping sweep method with DHCP reservation and DNS record check
            $script:resultListBox.Items.Add("Checking for potentially free IPs, DHCP reservations, and DNS records...")
            $startAddr = [System.Net.IPAddress]::Parse($startIP).GetAddressBytes()
            $endAddr = [System.Net.IPAddress]::Parse($endIP).GetAddressBytes()
            $freeCount = 0
            $ipsWithDnsRecords = @()
            
            for ($i = $startAddr[3]; $i -le $endAddr[3]; $i++) {
                $currentIP = "$($startAddr[0]).$($startAddr[1]).$($startAddr[2]).$i"
                $isPingable = Test-Connection -ComputerName $currentIP -Count 1 -Quiet
                $hasDhcpReservation = $false
                $dnsRecords = $null

                try {
                    $reservation = Get-DhcpServerv4Reservation -ComputerName $dhcpServer -IPAddress $currentIP -ErrorAction SilentlyContinue
                    $hasDhcpReservation = ($null -ne $reservation)
                } catch {
                    $errorMessage = $_.Exception.Message
                    $script:resultListBox.Items.Add("Error checking DHCP reservation for $currentIP`: $errorMessage")
                }

                try {
                    $dnsRecords = Resolve-DnsName -Name $currentIP -ErrorAction SilentlyContinue
                } catch {
                    # Do nothing, as no DNS record is expected for some IPs
                }
$dnsRecords = Resolve-DnsName -Name $currentIP -Server $dnsServer -ErrorAction SilentlyContinue
                if (!$isPingable -and !$hasDhcpReservation) {
                    if ($null -eq $dnsRecords) {
                        $script:resultListBox.Items.Add("Potentially free IP: $currentIP")
                        $freeCount++
                        if ($freeCount -ge 10) { break }
                    } else {
                        $ipsWithDnsRecords += [PSCustomObject]@{
                            IP = $currentIP
                            DNSRecords = $dnsRecords
                        }
                    }
                }
            }
            
            if ($freeCount -eq 0) {
                $script:resultListBox.Items.Add("No completely free IPs found. Showing IPs with DNS records but no DHCP reservation:")
                foreach ($ip in $ipsWithDnsRecords) {
                    $script:resultListBox.Items.Add("IP: $($ip.IP)")
                    foreach ($record in $ip.DNSRecords) {
                        $script:resultListBox.Items.Add("  DNS Record: $($record.Name) - $($record.Type)")
                    }
                }
            } else {
                $script:resultListBox.Items.Add("Search complete. Listed IPs are potentially free and have no DHCP reservations or DNS records.")
                $script:resultListBox.Items.Add("Note: Verify before use. There might be other network configurations to consider.")
            }
        } else {
            $errorMessage = "Please enter valid IP addresses for: " + ($invalidIPs -join " and ") + "`n`nUse the format: xxx.xxx.xxx.xxx"
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
}

        "Open GPO Packages" {
            $hintLabel.Text = "Instructions for Open GPO Packages:`n`n1. Click the 'Open GPO Packages' button to open the folder.`n`nHint: Ensure you have network access to the specified folder."
            
            $centerPane.Controls.Clear()
            
            # Create a panel to hold the content
            $contentPanel = New-Object System.Windows.Forms.Panel
            $contentPanel.Width = 400
            $contentPanel.Height = 200
            $contentPanel.Left = ($centerPane.ClientSize.Width - $contentPanel.Width) / 2
            $contentPanel.Top = ($centerPane.ClientSize.Height - $contentPanel.Height) / 2
            $centerPane.Controls.Add($contentPanel)

            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Open GPO Packages"
            $label.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
            $label.AutoSize = $true
            $label.Location = New-Object System.Drawing.Point(20, 20)

            $openButton = New-Object System.Windows.Forms.Button
            $openButton.Text = "Open GPO Packages Folder"
            $openButton.Location = New-Object System.Drawing.Point(20, 60)
            $openButton.Width = 200
            $openButton.Add_Click({
                $folderPath = "\\ds.ad.ssmhc.com\ssmdfs\GPO-Packages\SASH"
                if (Test-Path $folderPath) {
                    Start-Process explorer.exe -ArgumentList $folderPath
                } else {
                    [System.Windows.Forms.MessageBox]::Show("The specified folder path is not accessible. Please check your network connection and permissions.", "Folder Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            })

            $contentPanel.Controls.AddRange(@($label, $openButton))
        }
        
         "Open TC Printer Folder" {
            $hintLabel.Text = "Instructions Open TC Printer Folder:`n`n1. Click the 'Open TC Printer Folder' button to open the folder.`n`nHint: Ensure you have network access to the specified folder."
            
            $centerPane.Controls.Clear()
            
            # Create a panel to hold the content
            $contentPanel = New-Object System.Windows.Forms.Panel
            $contentPanel.Width = 400
            $contentPanel.Height = 200
            $contentPanel.Left = ($centerPane.ClientSize.Width - $contentPanel.Width) / 2
            $contentPanel.Top = ($centerPane.ClientSize.Height - $contentPanel.Height) / 2
            $centerPane.Controls.Add($contentPanel)

            $label = New-Object System.Windows.Forms.Label
            $label.Text = "Open TC Printer Folder"
            $label.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
            $label.AutoSize = $true
            $label.Location = New-Object System.Drawing.Point(20, 20)

            $openButton = New-Object System.Windows.Forms.Button
            $openButton.Text = "Open TC Printer Folder"
            $openButton.Location = New-Object System.Drawing.Point(20, 60)
            $openButton.Width = 200
            $openButton.Add_Click({
                $folderPath = "\\ds.ad.ssmhc.com\ssmdfs\GPO-Packages\SASH\Printers\User"
                if (Test-Path $folderPath) {
                    Start-Process explorer.exe -ArgumentList $folderPath
                } else {
                    [System.Windows.Forms.MessageBox]::Show("The specified folder path is not accessible. Please check your network connection and permissions.", "Folder Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                }
            })

            $contentPanel.Controls.AddRange(@($label, $openButton))
        }
    }
})

# Function to clone workstation
function Clone-Workstation {
    param($sourcePC, $targetPC)
    # Define the log file path
    $script:logDirectory = "C:\Temp\Powershell Scripts\Workstation Cloning"
    if (-not (Test-Path $script:logDirectory)) {
        New-Item -ItemType Directory -Path $script:logDirectory | Out-Null
    }
    $script:filePath = Join-Path $script:logDirectory "CloneWorkstation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

    # Define the Write-Log function within Clone-Workstation
    function Write-Log {
        param([string]$message)
        $message | Out-File -FilePath $script:filePath -Append
        Write-Host $message
    }

    Write-Log "Starting workstation cloning process at $(Get-Date)"
    Write-Log "Source PC: ${sourcePC}"
    Write-Log "Target PC: ${targetPC}"

    # Verify AD objects exist
    try {
        $sourceADObject = Get-ADComputer -Identity $sourcePC -Properties MemberOf -ErrorAction Stop
        $targetADObject = Get-ADComputer -Identity $targetPC -Properties MemberOf -ErrorAction Stop
    }
    catch {
        $errorMessage = "Error: One or both PCs not found. Please check the PC names and try again. Error: $($_.Exception.Message)"
        Write-Log $errorMessage
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "AD Object Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    Write-Log "Successfully found both source and target AD objects"

    # Get all AD group memberships for the source PC
    $sourceGroups = $sourceADObject.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
    
    Write-Log "Source PC group memberships:"
    $sourceGroups | ForEach-Object { Write-Log "  $_" }

    $skippedGroups = @()
    $addedGroups = @()

    Write-Log "Starting group membership cloning process"
    foreach ($group in $sourceGroups) {
        try {
            $targetCurrentGroups = $targetADObject.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
            if ($targetCurrentGroups -contains $group) {
                Write-Log "  Group ${group} is already a member of ${targetPC}, skipping"
                continue
            }

            Write-Log "  Attempting to add ${targetPC} to group ${group}"
            Add-ADGroupMember -Identity $group -Members $targetADObject -ErrorAction Stop
            
            # Verify the addition was successful
            $updatedTargetGroups = (Get-ADComputer -Identity $targetPC -Properties MemberOf).MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
            if ($updatedTargetGroups -contains $group) {
                $addedGroups += $group
                Write-Log "  Successfully added ${targetPC} to group ${group}"
            } else {
                throw "Group addition verification failed"
            }
        }
        catch {
            $skippedGroups += $group
            Write-Log "  Failed to add ${targetPC} to group ${group}. Error: $($_.Exception.Message)"
        }
    }

    # Update the ListBoxes
    $script:unableToAddBox.Items.Clear()
    $script:unableToAddBox.Items.AddRange($skippedGroups)
    $script:addedBox.Items.Clear()
    $script:addedBox.Items.AddRange($addedGroups)

    Write-Log "`nGroups Successfully Added:"
    $addedGroups | ForEach-Object { Write-Log "  $_" }

    Write-Log "`nGroups Unable to Add:"
    $skippedGroups | ForEach-Object { Write-Log "  $_" }

    # Verify final group memberships
    $finalTargetGroups = (Get-ADComputer -Identity $targetPC -Properties MemberOf).MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
    Write-Log "`nFinal group memberships for ${targetPC}:"
    $finalTargetGroups | ForEach-Object { Write-Log "  $_" }

    Write-Log "Workstation cloning process completed at $(Get-Date)"

    # Success message
    $successMessage = "Workstation cloning completed.`nAll actions were logged to: ${filePath}"
    [System.Windows.Forms.MessageBox]::Show($successMessage, "Clone Workstation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function Get-CurrentIPScope {
    # Get the IP address of the first active network adapter
    $ip = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.PrefixOrigin -eq 'Dhcp' } | Select-Object -First 1).IPAddress
    $subnet = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.PrefixOrigin -eq 'Dhcp' } | Select-Object -First 1).PrefixLength
    return "$ip/$subnet"
}

function Find-FreeIPsByPing {
    param (
        [string]$IPBase,
        [int]$Count = 20
    )

    $freeIPs = @()
    $thirdOctet = 0
    $fourthOctet = 1

    while ($freeIPs.Count -lt $Count -and $thirdOctet -le 255) {
        $currentIP = "${IPBase}.${thirdOctet}.${fourthOctet}"

        if (-not (Test-Connection -ComputerName $currentIP -Count 1 -Quiet)) {
            $freeIPs += $currentIP
        }

        $fourthOctet++
        if ($fourthOctet -gt 254) {
            $fourthOctet = 1
            $thirdOctet++
        }
    }

    return $freeIPs
}

function Test-FreeIP {
    param (
        [string]$IP,
        [string]$DHCPServer,
        [string]$DNSServer,
        [array]$DHCPScopes
    )

    # Check if IP responds to ping
    if (Test-Connection -ComputerName $IP -Count 1 -Quiet) {
        return $false
    }

    # Check DHCP reservation and lease
    try {
        $dhcpScope = $DHCPScopes | Where-Object { $IP -ge $_.StartRange -and $IP -le $_.EndRange } | Select-Object -First 1
        if ($dhcpScope) {
            $dhcpReservation = Get-DhcpServerv4Reservation -ComputerName $DHCPServer -ScopeId $dhcpScope.ScopeId -IPAddress $IP -ErrorAction SilentlyContinue
            $dhcpLease = Get-DhcpServerv4Lease -ComputerName $DHCPServer -ScopeId $dhcpScope.ScopeId -IPAddress $IP -ErrorAction SilentlyContinue
            if ($dhcpReservation -or $dhcpLease) {
                return $false
            }
        }
    }
    catch {
        Write-Host "Error checking DHCP: $_"
    }

    # Check DNS record
    try {
        $dnsRecord = Resolve-DnsName -Name $IP -Server $DNSServer -ErrorAction SilentlyContinue
        if ($dnsRecord) {
            return $false
        }
    }
    catch {
        Write-Host "Error checking DNS: $_"
    }

    return $true
}



# Show the form
$form.ShowDialog()
