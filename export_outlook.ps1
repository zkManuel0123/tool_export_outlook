# Outlook Calendar Export Script
# Requires Outlook installed and configured

# Export settings
$outputPath = "D:\outlookExport\OutlookCalendar_Export.txt"
$daysAhead = 30
$daysBack = 7

# Pause function to prevent window from closing quickly
function Pause-Script {
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor Cyan
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

try {
    # Check and set execution policy
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
    if ($currentPolicy -eq "Restricted") {
        Write-Host "Setting PowerShell execution policy..." -ForegroundColor Yellow
        try {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
            Write-Host "Execution policy set successfully!" -ForegroundColor Green
        } catch {
            Write-Host "Cannot set execution policy, please run PowerShell as administrator" -ForegroundColor Red
            Pause-Script
            exit 1
        }
    }
    
    # Create Outlook application object
    Write-Host "Connecting to Outlook..." -ForegroundColor Green
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    # Get default calendar folder
    $calendar = $namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
    
    # Set time range
    $startDate = (Get-Date).AddDays(-$daysBack).ToString("MM/dd/yyyy 00:00")
    $endDate = (Get-Date).AddDays($daysAhead).ToString("MM/dd/yyyy 23:59")
    
    Write-Host "Export time range: $startDate to $endDate" -ForegroundColor Yellow
    
    # Filter calendar items
    $filter = "[Start] >= '$startDate' AND [Start] <= '$endDate'"
    $appointments = $calendar.Items.Restrict($filter)
    $appointments.Sort("[Start]")
    
    # Create output content
    $output = @()
    $output += "=" * 60
    $output += "Outlook Calendar Export Report"
    $output += "Export time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $output += "Time range: $startDate to $endDate"
    $output += "Found $($appointments.Count) appointments"
    $output += "=" * 60
    $output += ""
    
    if ($appointments.Count -eq 0) {
        $output += "No appointments found in the specified time range."
    } else {
        $currentDate = ""
        
        foreach ($appointment in $appointments) {
            # Format date
            $appointmentDate = $appointment.Start.ToString("yyyy-MM-dd")
            $appointmentTime = $appointment.Start.ToString("HH:mm")
            $endTime = $appointment.End.ToString("HH:mm")
            
            # Add date separator for new day
            if ($appointmentDate -ne $currentDate) {
                $output += ""
                $output += "[ $appointmentDate ]"
                $output += "-" * 30
                $currentDate = $appointmentDate
            }
            
            # Add meeting details
            $duration = ""
            if ($appointment.AllDayEvent) {
                $duration = "All Day"
            } else {
                $duration = "$appointmentTime - $endTime"
            }
            
            $output += "Time: $duration"
            $output += "Subject: $($appointment.Subject)"
            
            if ($appointment.Location -and $appointment.Location.Trim() -ne "") {
                $output += "Location: $($appointment.Location)"
            }
            
            if ($appointment.Organizer -and $appointment.Organizer.Trim() -ne "") {
                $output += "Organizer: $($appointment.Organizer)"
            }
            
            # Get attendees information
            if ($appointment.Recipients.Count -gt 0) {
                $attendees = @()
                foreach ($recipient in $appointment.Recipients) {
                    $attendees += $recipient.Name
                }
                if ($attendees.Count -gt 0) {
                    $output += "Attendees: $($attendees -join ', ')"
                }
            }
            
            if ($appointment.Body -and $appointment.Body.Trim() -ne "") {
                $bodyPreview = $appointment.Body.Substring(0, [Math]::Min(100, $appointment.Body.Length)).Replace("`n", " ").Replace("`r", "")
                $output += "Notes: $bodyPreview$(if($appointment.Body.Length -gt 100){'...'})"
            }
            
            $output += ""
        }
    }
    
    $output += "=" * 60
    $output += "Export completed!"
    
    # Write to file
    $output | Out-File -FilePath $outputPath -Encoding UTF8
    
    Write-Host "Export successful!" -ForegroundColor Green
    Write-Host "File saved to: $outputPath" -ForegroundColor Cyan
    Write-Host "Exported $($appointments.Count) appointments" -ForegroundColor Yellow
    
    # Ask if user wants to open the file
    $openFile = Read-Host "Open the exported file? (Y/N)"
    if ($openFile -eq "Y" -or $openFile -eq "y") {
        Start-Process notepad.exe $outputPath
    }
    
    Pause-Script
    
} catch {
    Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please ensure:" -ForegroundColor Yellow
    Write-Host "1. Outlook is installed and running" -ForegroundColor Yellow
    Write-Host "2. Email account is configured" -ForegroundColor Yellow
    Write-Host "3. Run PowerShell as administrator" -ForegroundColor Yellow
    
    Pause-Script
} finally {
    # Clean up COM objects
    if ($outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}