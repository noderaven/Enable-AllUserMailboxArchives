#Requires -Module ExchangeOnlineManagement

<#
.SYNOPSIS
    Enables online archives for all user mailboxes in Microsoft 365 that don't already have one.

.DESCRIPTION
    This script connects to Exchange Online, retrieves a list of all regular user mailboxes,
    filters out any that already have an archive enabled, and then enables the
    archive for the remaining mailboxes. It specifically excludes shared, room,
    equipment, and other non-user mailbox types by only targeting 'UserMailbox' recipients.
    The script provides a final summary of all users that were successfully modified.

.NOTES
    Prerequisites:
    - The ExchangeOnlineManagement PowerShell module must be installed.
      Run in PowerShell as Administrator: Install-Module -Name ExchangeOnlineManagement -Force
    - You must have the appropriate permissions in Exchange Online (e.g., Exchange Administrator role)
      to connect and modify mailbox properties.
#>

# --- Script Body ---
try {
    # Announce the start of the script
    Write-Host "Starting script to enable online archives for user mailboxes..." -ForegroundColor Cyan

    # --- Step 1: Connect to Exchange Online ---
    # Check if a session is already active. If not, establish a new connection.
    # You will be prompted for credentials if not already authenticated.
    # Use -UserPrincipalName youradmin@domain.com for non-interactive sign-in if needed.
    $currentSession = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if (-not $currentSession) {
        Write-Host "Connecting to Exchange Online PowerShell..."
        Connect-ExchangeOnline -ShowBanner:$false
    } else {
        Write-Host "Already connected to Exchange Online (Organization: $($currentSession.Organization))"
    }

    # --- Step 2: Get User Mailboxes Without an Archive ---
    # Retrieve all mailboxes where the recipient type is 'UserMailbox' and an archive has not been provisioned.
    # Using a string-based -Filter is the most reliable method.
    # A '000...' GUID indicates no archive exists.
    Write-Host "Searching for user mailboxes without an active online archive. This may take a moment..."
    $mailboxesToEnable = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Filter "ArchiveGuid -eq '00000000-0000-0000-0000-000000000000'" -ErrorAction Stop
    
    # Create a list to hold the names of users that are successfully modified
    $modifiedUsers = [System.Collections.Generic.List[string]]::new()

    # --- Step 3: Enable Archives and Report Progress ---
    if ($null -eq $mailboxesToEnable -or $mailboxesToEnable.Count -eq 0) {
        Write-Host "Success: All user mailboxes already have an online archive enabled." -ForegroundColor Green
    }
    else {
        # Get a count for the progress indicator
        $total = ($mailboxesToEnable | Measure-Object).Count
        $count = 0
        Write-Host "Found $total user mailbox(es) that require an archive." -ForegroundColor Yellow

        # Loop through each identified mailbox and enable its archive
        foreach ($mailbox in $mailboxesToEnable) {
            $count++
            $upn = $mailbox.UserPrincipalName
            Write-Host "($count/$total) Processing mailbox: $upn"

            try {
                # CORRECTED: Use Enable-Mailbox with the -Archive switch to enable the archive
                Enable-Mailbox -Identity $upn -Archive -ErrorAction Stop
                Write-Host "  -> Successfully enabled archive for $upn" -ForegroundColor Green
                # Add the user to our list of modified mailboxes for the final summary
                $modifiedUsers.Add($upn)
            }
            catch {
                # Report any errors encountered for a specific mailbox
                # CORRECTED: Using -f format operator for safer string construction
                Write-Warning ("  -> FAILED to enable archive for {0}. Error: {1}" -f $upn, $_.Exception.Message)
            }
        }
    }

    # --- Step 4: Final Summary ---
    Write-Host "`nScript execution complete." -ForegroundColor Cyan
    
    # Report on which users were actually modified
    if ($modifiedUsers.Count -gt 0) {
        # CORRECTED: Fixed invalid characters and quotes in the string
        Write-Host "`nSummary of successfully enabled archives ($($modifiedUsers.Count)):" -ForegroundColor Cyan
        foreach ($user in $modifiedUsers) {
            Write-Host " - $user"
        }
    }

    Write-Host "`nNote: It may take some time for newly enabled archives to appear for users in Outlook and Outlook on the web."

}
catch {
    # Catch any script-terminating errors (e.g., connection failure, permission issues)
    # CORRECTED: Using -f format operator for safer string construction
    Write-Error ("A critical error occurred: {0}" -f $_.Exception.Message)
}
finally {
    # --- Step 5: Disconnect Session (Optional) ---
    # It's good practice to disconnect your session when finished.
    # Uncomment the line below if you want the script to automatically disconnect.
    #
    # Disconnect-ExchangeOnline -Confirm:$false
}
