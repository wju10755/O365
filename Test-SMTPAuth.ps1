<#
.SYNOPSIS
A powershell tool to validated functionality of SMTP Authentication using the default Office 365 SMTP servers or by specifying a custom SMTP server.

.DESCRIPTION
This tool allows engineers to validate the functionality of the SMTP authentication. 
You are required to provide the following information: Sender Address, Password, Recipient address, SMTP Server Address, SMTP Protocol, and Port Number. 
Additional field requirement details can be found under the Help>About menu.

.PARAMETER None

.INPUTS
Sender Address - Full email address of the sender.
Password - Credentials for sending account.
Recpient Address - Recepient of the test message.

SMTP Server: Select from standard Office 365 SMTP server (smtp.office365.com), legacy Office 365 SMTP server (smtp-legacy.office365.com), or provide a custom SMTP server in the pop-up window.

.EXAMPLE
.\Test-SMTPAuth.ps1

.NOTES
Author: [Bill Ulrich]
Date: [6/25/2024]
Version: 7.0

#>
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'MITS - SMTP Authentication Test'
$form.Size = New-Object System.Drawing.Size(300,200)

# Hide the console window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();

[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);

public static void HideConsoleWindow()
{
    var handle = GetConsoleWindow();
    ShowWindow(handle, 0); // 0 = SW_HIDE
}'

[Console.Window]::HideConsoleWindow()



function IsValidEmail($email) {
    return $email -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
}

$ProgressPreference = "SilentlyContinue"
$url1 = "https://advancestuff.hostedrmm.com/labtech/transfer/installers/labtech.ico"
$destinationPath1 = "$env:Temp\labtech.ico"
Invoke-WebRequest -Uri $url1 -OutFile $destinationPath1 -ErrorAction SilentlyContinue > $null 2>&1

$url2 = "https://advancestuff.hostedrmm.com/labtech/transfer/installers/redA.png"
$destinationPath2 = "$env:Temp\redA.png"
Invoke-WebRequest -Uri $url2 -OutFile $destinationPath2 -ErrorAction SilentlyContinue > $null 2>&1
$ProgressPreference = 'Continue'

$form = New-Object System.Windows.Forms.Form
$form.Text = 'MITS - SMTP Authentication Test'
$form.Size = New-Object System.Drawing.Size(500, 350)
$form.StartPosition = 'CenterScreen'

# Set the form icon
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($destinationPath1)

#Prevent resizing
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

# Optionally, hide maximize and minimize buttons
$form.MaximizeBox = $false
$form.MinimizeBox = $false

# Create MenuStrip
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$fileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem 'File'
$clearFieldsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem 'Clear Fields'
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem 'Exit'

# Add items to MenuStrip
$fileMenuItem.DropDownItems.Add($clearFieldsMenuItem)
$fileMenuItem.DropDownItems.Add($exitMenuItem)
$menuStrip.Items.Add($fileMenuItem)
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

# Create Help Menu Item
$helpMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem 'Help'
$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem 'About'

# Add About to Help Menu
$helpMenuItem.DropDownItems.Add($aboutMenuItem)

# Add Help Menu to MenuStrip
$menuStrip.Items.Add($helpMenuItem)

$aboutMenuItem.Add_Click({
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = 'About'
    $aboutForm.Size = New-Object System.Drawing.Size(400, 320)
    $aboutForm.StartPosition = 'CenterScreen'
    $aboutForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

    # Set the icon for the About form
    #$aboutForm.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon('$env:Temp\labtech.ico')
    
    # Create the first label for the bold text
    $boldLabel = New-Object System.Windows.Forms.Label
    $boldLabel.Location = New-Object System.Drawing.Point(10, 10) # Adjust location as needed
    $boldLabel.Size = New-Object System.Drawing.Size(240, 30) # Adjusted size to prevent overlap with the image
    $boldLabel.Text = "Advance Managed IT"
    $boldLabel.Font = New-Object System.Drawing.Font($boldLabel.Font.FontFamily, 15, [System.Drawing.FontStyle]::Bold)
    
    # Create the mits label for the additional text
    $mitsLabel = New-Object System.Windows.Forms.Label
    $mitsLabel.Location = New-Object System.Drawing.Point(10, 55) # Placed underneath the first label
    $mitsLabel.Size = New-Object System.Drawing.Size(290, 30) # Adjust size as needed
    $mitsLabel.Text = "MITS - SMTP Authentication Tester"
    $mitsLabel.Font = New-Object System.Drawing.Font($mitsLabel.Font.FontFamily, 11) # Adjust font size as needed
    $aboutForm.Controls.Add($mitsLabel)

    # Create an 'OK' button for the form
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = 'OK'
     # Adjust location as needed
    $okButton.Location = New-Object System.Drawing.Point(140, 245)    
    $okButton.Size = New-Object System.Drawing.Size(100, 30)
    $okButton.Add_Click({ $aboutForm.Close() })

    # Add the PictureBox to display the image
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Location = New-Object System.Drawing.Point(305, 09) 
    $pictureBox.Size = New-Object System.Drawing.Size(100, 100)  
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::AutoSize
    $pictureBox.Image = [System.Drawing.Image]::FromFile($destinationPath2)

    # Create the input box (TextBox) under the image
    $inputBox = New-Object System.Windows.Forms.TextBox
    $inputBox.Location = New-Object System.Drawing.Point(0, 100)  
    # Assuming $form is your form object
    $inputBox.Size = New-Object System.Drawing.Size($form.ClientSize.Width, 130) 
    #$inputBox.Size = New-Object System.Drawing.Size($form.ClientSize.Height,10, 150)
    $inputBox.Multiline = $true # Enable multiline input
    $inputBox.BackColor = [System.Drawing.Color]::Black  
    $inputBox.ForeColor = [System.Drawing.Color]::White  
    $inputBox.Text = "`nSender Address: Email address of the sender.`r`n`r`nPassword: Password for the sending account.`r`n`r`nRecipient Address: Email address of the recipient.`r`n`r`nSMTP Server: Select between standard, legacy, or other.`r`n`r`nPort Number: Automatically updates when a protocol is selected."
    # Add controls to the form
    $aboutForm.Controls.Add($okButton)
    $aboutForm.Controls.Add($boldLabel)
    $aboutForm.Controls.Add($pictureBox)
    $aboutForm.Controls.Add($inputBox) # Add the input box to the form
    
    $aboutForm.Add_Resize({
        $inputBox.Width = $aboutForm.ClientSize.Width
        $inputBox.Height = $aboutForm.ClientSize.Height - $inputBox.Location.Y - 1 # Adjust the 20 to leave some margin at the bottom
        # Adjust okButton size and location
    $okButton.Size = New-Object System.Drawing.Size(100, 30) # You can make the size dynamic based on form size if needed
    $okButton.Location = New-Object System.Drawing.Point(($aboutForm.ClientSize.Width - $okButton.Width) / 2, $aboutForm.ClientSize.Height - $okButton.Height - 20) # Center the button and place it 20 pixels above the bottom of the form
    })

    # Show the About form as a dialog
    $aboutForm.ShowDialog()
})

# Copier Email Address
$labelCopier = New-Object System.Windows.Forms.Label
$labelCopier.Location = New-Object System.Drawing.Point(10, 40)
$labelCopier.Size = New-Object System.Drawing.Size(160, 20)
$labelCopier.Text = 'Sender Address:'
$form.Controls.Add($labelCopier)

$textboxCopier = New-Object System.Windows.Forms.TextBox
$textboxCopier.Location = New-Object System.Drawing.Point(180, 40)
$textboxCopier.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($textboxCopier)

# Password for Sending Address
$labelPassword = New-Object System.Windows.Forms.Label
$labelPassword.Location = New-Object System.Drawing.Point(10, 70)
$labelPassword.Size = New-Object System.Drawing.Size(160, 20)
$labelPassword.Text = 'Password:'
$form.Controls.Add($labelPassword)

$textboxPassword = New-Object System.Windows.Forms.TextBox
$textboxPassword.Location = New-Object System.Drawing.Point(180, 70)
$textboxPassword.Size = New-Object System.Drawing.Size(300, 20)
$textboxPassword.UseSystemPasswordChar = $true
$form.Controls.Add($textboxPassword)

# Recipient Email Address
$labelRecipient = New-Object System.Windows.Forms.Label
$labelRecipient.Location = New-Object System.Drawing.Point(10, 100)
$labelRecipient.Size = New-Object System.Drawing.Size(160, 20)
$labelRecipient.Text = 'Recipient Address:'
$form.Controls.Add($labelRecipient)

$textboxRecipient = New-Object System.Windows.Forms.TextBox
$textboxRecipient.Location = New-Object System.Drawing.Point(180, 100)
$textboxRecipient.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($textboxRecipient)

# SMTP Server Address - Modify to ComboBox
$labelSMTPSvr = New-Object System.Windows.Forms.Label
$labelSMTPSvr.Location = New-Object System.Drawing.Point(10, 130)
$labelSMTPSvr.Size = New-Object System.Drawing.Size(160, 20)
$labelSMTPSvr.Text = 'SMTP Server Address:'
$form.Controls.Add($labelSMTPSvr)
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($destinationPath1)

$comboSMTPSvr = New-Object System.Windows.Forms.ComboBox
$comboSMTPSvr.Location = New-Object System.Drawing.Point(180, 130)
$comboSMTPSvr.Size = New-Object System.Drawing.Size(300, 20)
$comboSMTPSvr.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboSMTPSvr.Items.AddRange(@('smtp.office365.com', 'smtp-legacy.office365.com', 'other'))
$comboSMTPSvr.SelectedIndex = 0 # Pre-select 'smtp.office365.com'

$form.Controls.Add($comboSMTPSvr)

# SMTP Protocol Selection
$labelSMTPProtocol = New-Object System.Windows.Forms.Label
$labelSMTPProtocol.Location = New-Object System.Drawing.Point(10, 160)
$labelSMTPProtocol.Size = New-Object System.Drawing.Size(160, 20)
$labelSMTPProtocol.Text = 'SMTP Protocol:'
$form.Controls.Add($labelSMTPProtocol)

$comboSMTPProtocol = New-Object System.Windows.Forms.ComboBox
$comboSMTPProtocol.Location = New-Object System.Drawing.Point(180, 160)
$comboSMTPProtocol.Size = New-Object System.Drawing.Size(100, 20)
$comboSMTPProtocol.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboSMTPProtocol.Items.AddRange(@('TLS', 'SSL', 'DirectSend'))
$form.Controls.Add($comboSMTPProtocol)
$comboSMTPProtocol.SelectedIndex = 0 # Pre-select 'SSL'

$comboSMTPProtocol.Add_SelectedIndexChanged({
    if ($comboSMTPProtocol.SelectedItem -eq 'TLS') {
        $textboxPortNumber.Text = '587'
    } elseif ($comboSMTPProtocol.SelectedItem -eq 'SSL') {
        $textboxPortNumber.Text = '465'
    } elseif ($comboSMTPProtocol.SelectedItem -eq 'DirectSend') {
        $textboxPortNumber.Text = '25'
    }
})
# Add 'DirectSend' to the SMTP Protocol ComboBox
#$comboSMTPProtocol.Items.Add('DirectSend')

# Port Number Input
$labelPortNumber = New-Object System.Windows.Forms.Label
$labelPortNumber.Location = New-Object System.Drawing.Point(290, 160)
$labelPortNumber.Size = New-Object System.Drawing.Size(80, 20)
$labelPortNumber.Text = 'Port Number:'
$form.Controls.Add($labelPortNumber)

$textboxPortNumber = New-Object System.Windows.Forms.TextBox
$textboxPortNumber.Location = New-Object System.Drawing.Point(370, 160)
$textboxPortNumber.Size = New-Object System.Drawing.Size(110, 20)
$textboxPortNumber.Text = '587' # Pre-set the default port number to '587'
$form.Controls.Add($textboxPortNumber)

# Output Box
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(10, 240)
$outputBox.Size = New-Object System.Drawing.Size(470, 60)
$outputBox.Multiline = $true
$outputBox.ReadOnly = $true
$outputBox.ScrollBars = 'Vertical'
$outputBox.BackColor = [System.Drawing.Color]::Black
$outputBox.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($outputBox)

# Clear Button
$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Location = New-Object System.Drawing.Point(290, 200) # Position to the right of the 'Send' button
$buttonClear.Size = New-Object System.Drawing.Size(100, 30)
$buttonClear.Text = 'Clear'
$form.Controls.Add($buttonClear)

$buttonClear.Add_Click({
    $outputBox.Clear() # Clear the output box when 'Clear' button is clicked
})


# Send Button
$buttonSend = New-Object System.Windows.Forms.Button
$buttonSend.Location = New-Object System.Drawing.Point(180, 200)
$buttonSend.Size = New-Object System.Drawing.Size(100, 30)
$buttonSend.Text = 'Send'
$form.Controls.Add($buttonSend)

# Add 'other' option to the SMTP Server Address combo box


# Modify the Send Button Click Event to use the selected SMTP Server Address
$buttonSend.Add_Click({
    $Copier = $textboxCopier.Text
    $Recipient = $textboxRecipient.Text
    $SMTPSvr = $comboSMTPSvr.SelectedItem # Use the selected item from the combo box
    if ($SMTPSvr -eq 'other') {
        # Use System.Windows.Forms to prompt for the 'other' SMTP server address
        $prompt = New-Object System.Windows.Forms.Form
        $prompt.Text = 'Enter SMTP Server Address'
        $prompt.Size = New-Object System.Drawing.Size(300, 150)
        $prompt.StartPosition = 'CenterScreen'
        $prompt.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($destinationPath1)

        $label = New-Object System.Windows.Forms.Label
        $label.Text = 'SMTP Server Address:'
        $label.Location = New-Object System.Drawing.Point(10, 20)
        $label.Size = New-Object System.Drawing.Size(280, 20)
        $prompt.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10, 40)
        $textBox.Size = New-Object System.Drawing.Size(260, 20)
        $prompt.Controls.Add($textBox)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = 'OK'
        $okButton.Location = New-Object System.Drawing.Point(10, 70)
        $okButton.Size = New-Object System.Drawing.Size(75, 23)
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $prompt.Controls.Add($okButton)
        $prompt.AcceptButton = $okButton

        $result = $prompt.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $SMTPSvr = $textBox.Text
        } else {
            $outputBox.AppendText("SMTP Server Address is required.`n")
            return
        }
    }
    $Password = $textboxPassword.Text
    $Port = $textboxPortNumber.Text
    $Protocol = $comboSMTPProtocol.SelectedItem

    $outputBox.Clear() # Clear the output box at the beginning of each send attempt
    $outputBox.AppendText("Sending test message...`r`n`r`n") # Display sending message

    if ($Protocol -ne 'DirectSend') {
        if (-not (IsValidEmail $Copier)) {
            $outputBox.AppendText("The copier email address is not valid.`n")
            return
        }
    }

    if (-not (IsValidEmail $Recipient)) {
        $outputBox.AppendText("The recipient email address is not valid.`n")
        return
    }

    if ($Protocol -eq 'DirectSend') {
        try {
            # For DirectSend, no authentication is required, so omit -Credential and -UseSsl parameters
            Send-MailMessage -From $Copier -To $Recipient -Subject "Test Email" -Body "Test SMTP Service from PowerShell using DirectSend." -SmtpServer $SMTPSvr -Port $Port -ErrorAction Stop
            $outputBox.AppendText("Email sent successfully!`n")
        } catch {
            $outputBox.AppendText("Failed to send email. Error: $($_.Exception.Message)`n")
        }
    } else {
        $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
        $Creds = New-Object System.Management.Automation.PSCredential ($Copier, $SecurePassword)
        
        try {
            Send-MailMessage -From $Copier -To $Recipient -Subject "Test Email" -Body "Test SMTP Service from PowerShell on Port $Port with $Protocol." -SmtpServer $SMTPSvr -Credential $Creds -UseSsl -Port $Port -ErrorAction Stop
            $outputBox.AppendText("Email sent successfully!`n")
        } catch {
            $outputBox.AppendText("Failed to send email. Error: $($_.Exception.Message)`n")
        }
    }
})

# Clear Fields Menu Item Click Event
$clearFieldsMenuItem.Add_Click({
    $textboxCopier.Clear()
    $textboxPassword.Clear()
    $textboxRecipient.Clear()
    $comboSMTPSvr.SelectedIndex = -1
    $textboxPortNumber.Clear()
    $comboSMTPProtocol.SelectedIndex = -1
    $outputBox.Clear()
})

# Exit Menu Item Click Event
$exitMenuItem.Add_Click({
    $form.Close()
})

# Adjust form controls to accommodate the MenuStrip, moving all except the output box up one line
$form.Controls | ForEach-Object { 
    if ($_.Top -gt $menuStrip.Height - 20 -and $_ -ne $outputBox) { # Exclude the output box and adjust for one line up
        $_.Location = New-Object System.Drawing.Point($_.Location.X, ($_.Location.Y + $menuStrip.Height - 20)) 
    }
}
$form.Height += $menuStrip.Height - 20 # Adjust form height accordingly

$form.ShowDialog()
