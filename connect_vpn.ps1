param (
    [string]$vpnName,
    [string]$phonebookPath
)

# Start the rasphone.exe process with the specified phonebook and VPN name
$arguments = "-f `"$phonebookPath`" -d `"$vpnName`""
Write-Output "Starting rasphone.exe with arguments: $arguments"
Start-Process "rasphone.exe" -ArgumentList $arguments

# Wait for the window to appear
Write-Output "Waiting for the VPN window to appear..."
Start-Sleep -Seconds 5

# Add the necessary assembly for SendKeys
Add-Type -AssemblyName System.Windows.Forms

# Function to send a key press to the foreground window
function Send-Enter {
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
}

# Send the ENTER key to proceed with the VPN connection
Write-Output "Sending first ENTER key..."
Send-Enter

# Wait for the username/password window to appear
Write-Output "Waiting for the username/password window..."
Start-Sleep -Seconds 5

# Send the ENTER key again to proceed with the VPN connection
Write-Output "Sending second ENTER key..."
Send-Enter

Write-Output "VPN connection process completed."
