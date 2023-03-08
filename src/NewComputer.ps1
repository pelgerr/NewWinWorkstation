# Script for adding new computers to a Windows domain. Insitutional use.
# Author: Rob Pelger
# Revised: 2022.05.06
# Revisions: Add Windows.Forms objects

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Initialize variables for safe memory allocation
$hostname = "null"
$description = "null"
$campus = "null"
$counter = 0

### FUNCTION DEFINITIONS
# New-Window creates the basic window for receiving user input and displaying output
# Takes three arguments: $windowTitle, $bodyText, $bodyType
# $bodyType options are: 0 - Text field, 1 - List, 2 - Print to user
function New-Window{
	param ([string]$windowTitle, [string]$bodyText, [int]$bodyType)
	# Create window
	$form = New-Object System.Windows.Forms.Form
	$form.Text = $windowTitle
	$form.Size = New-Object System.Drawing.Size(300,200)
	$form.StartPosition = 'CenterScreen'

	# OK button
	$okButton = New-Object System.Windows.Forms.Button
	$okButton.Location = New-Object System.Drawing.Point(75,120)
	$okButton.Size = New-Object System.Drawing.Size(75,23)
	$okButton.Text = 'OK'
	$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form.AcceptButton = $okButton
	$form.Controls.Add($okButton)

	# Cancel button
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton.Location = New-Object System.Drawing.Point(150,120)
	$cancelButton.Size = New-Object System.Drawing.Size(75,23)
	$cancelButton.Text = 'Cancel'
	$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form.CancelButton = $cancelButton
	$form.Controls.Add($cancelButton)

	# Body text
	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Point(10,20)
	$label.Size = New-Object System.Drawing.Size(280,30)
	$label.Text = $bodyText
	$form.Controls.Add($label)

	# Create text field, list field, or print info (no input field)
	Switch ($bodyType) {
		# 0 - Text field
		0 {
			$textBox = New-Object System.Windows.Forms.TextBox
			$textBox.Location = New-Object System.Drawing.Point(10,50)
			$textBox.Size = New-Object System.Drawing.Size(260,20)

			$form.Controls.Add($textBox)
			$form.Topmost = $true
			$form.Add_Shown({$textBox.Select()})
			$result = $form.ShowDialog()
			$returnText = $textBox.Text
		}
		# 1 - List box
		1 {
			$listBox = New-Object System.Windows.Forms.ListBox
			$listBox.Location = New-Object System.Drawing.Point(10,50)
			$listBox.Size = New-Object System.Drawing.Size(260,20)
			$listBox.Height = 80

			# List options
			[void] $listBox.Items.Add('LOC1')
			[void] $listBox.Items.Add('LOC2')
			[void] $listBox.Items.Add('LOC3')
			[void] $listBox.Items.Add('LOC4')

			$form.Controls.Add($listBox)
			$form.Topmost = $true
			$result = $form.ShowDialog()
			$returnText = $listBox.SelectedItem
		} 
		# 2 - Print to user
		2 {
			$label.Size = New-Object System.Drawing.Size(280,100)
			$label.Text = $bodyText
			$form.Controls.Add($label)
			$form.Topmost = $true
			$result = $form.ShowDialog()		
			$returnText = $result
		}
	} 
	# If OK button is clicked, proceed. Otherwise, exit script.
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
		return $returnText
	}
	else {
		Write-Output "Transaction aborted. Please run the script again."
		Exit
	}
}

# Get-Description will take user input ($desc) and determine which description suffix to concatenate to ($desc)
# based on the last character in ($hostname). It will then create a local system object and update the computer description
# using ($desc)
function Get-Description {
	$windowTitle = 'DESCRIPTION'
	$bodyText = 'Please enter the name of the user or department in the space below:'
	$bodyType = 0

	$desc = New-Window $windowTitle $bodyText $bodyType
	
	# Append machine type to description
    if ($hostname -match 'L$') {
        $desc += "'s Laptop"
    }
    else {
        $desc += "'s Desktop"
    }

	# Write description changes to local machine
    $localObject = Get-WMIObject Win32_OperatingSystem
    $localObject.Description = $desc
	# DEBUG
    #$localObject.Put()
    return $desc
}

# Push-ADChanges will prompt for domain credentials and add the computer to the correct OU in AD, rename the computer to ($hostname),
# and update the description of the computer in AD to ($description)
function Push-ADChanges {
    # Get user domain credentials
	# DEBUG
    #$credential = Get-Credential
    # Add computer to the correct OU, change hostname
	# DEBUG
    #Add-Computer -DomainName "GordonConwell.edu" -OUPath OU=$campus",OU=Windows,OU=Workstations,OU=Hardware,OU=1GCTS,DC=gordonconwell,DC=edu" -Credential $credential -NewName $hostname
    ## BROKEN - Needs ActiveDirectory module installed
    ## TODO: Either get AD module installed or find different way to push description changes
    ## Push description change to AD
    ## Set-ADComputer -Identity $hostname -Description $description -Credential $credential
	# DEBUG
    #Write-Output "Active Directory successfully updated! Restarting..."
}

### BEGIN MAIN
# HOSTNAME LOOP
# TODO: Expand validation to follow specific naming convention
# Prompt user for hostname and set loop exit condition
while ($counter -ne 1) {
	$title = 'HOSTNAME'
	$body = 'Please enter the name of the computer:'
	$bodyType = 0
	$hostname = New-Window $title $body $bodyType 
	$hostname = $hostname.ToUpper()
    # Validate user input for correct hostname length: XXX#####-XX (11 chars)
    if ($hostname.length -eq 11) {
    # Increase counter to meet exit condition
        $counter++
    }
    else {
        Write-Output "`nIncorrect naming convention. Please follow the guidelines in the hardware naming convention KBA."
    }
}

# CAMPUS 
# Select the registered campus
$title = 'CAMPUS'
$body = 'Please select the campus:'
$bodyType = 1
$campus = New-Window $title $body $bodyType

# DESCRIPTION PROCESSING 
# Call Get-Description and store output in $description casted to string data type
# Remove leading path data that gets added after receiving $desc from Get-Description, using Split()
# If no path data is found, Split() will return the contents of $description
[string]$description = Get-Description
$description = $description.Split('@')[-1]

# SUMMARY
# Print summary of changes to be pushed
$title = 'SUMMARY'
$body = "The following changes will be made: `n`n Computer name: $hostname `n Registered campus: $campus `n Computer description: $description `n"
$bodyType = 2
$result = New-Window $title $body $bodyType

# Push changes, restrict ExecutionPolicy, and reboot if OK button is pressed. Otherwise, abort script and exit
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
	# DEBUG
    #Push-ADChanges
    # Restrict script execution policy before exiting
	# DEBUG
    #Set-ExecutionPolicy restricted
	# DEBUG
    #Restart-Computer
}
else {
    Write-Output "Transaction aborted. Please run the script again."
    Exit
}

################################################
# TODO: Either get AD module installed or find different way to push description changes
# TODO: Add detailed errors and error handling
# TODO: Sanitize user input
################################################
# HOW TO USE THIS SCRIPT
# 1) Open PowerShell as an Administrator
# 2) Run: Set-ExecutionPolicy unrestricted
# 3) Type 'a' then hit Enter
# 4) Type (or copy from here and paste) cd C:\Temp\
# 5) Hit Enter
# 6) Type .\NewComputer.ps1
# 7) Hit Enter
# 8) Follow the prompts
