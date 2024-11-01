Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Users.Actions
Import-Module ExchangeOnlineManagement


# Loops Indefinitely Until User Selects 0 To Exit
while ($true) {

    try {
        # Prompts For User Selection
        $userChoice = Read-Host "Enter your number selection`n0)Exit :`n1)Scan For Inbox Rules : `n2)Scan For Disabled User Accounts [Select User(s) To Retrieve License Information] : `n3)Convert To Shared Mailbox [Select User(s) To Convert Mailbox] : `n4)Convert Back To Regular Mailbox [Select User(s) To Convert Mailbox] : `n5)Revoke User(s) Session :`n=>Selection "

        # Provides A Retry If User Has Selected Alpha Characters
        if (-not [int]::TryParse($userChoice, [ref]$null)){
            throw "Invalid Input."
        }

        # Exits The Script If Zero Is Typed Out
        if ($userChoice -eq 0) {
            # Exit the loop if user selects 0
            Write-Host -ForegroundColor Yellow "Exiting..."
            break
        }
        
        # Code Block For Inbox Rule
        if ($userChoice -eq 1) {
            try{
                 # Automatically signs the user for Exchange Online Server
                Connect-ExchangeOnline -UserPrincipalName upn_here  -Organization "domain.xyz" -ErrorAction Stop

                # Delay added so script doesn't run everything all at once
                Write-Host -ForegroundColor Yellow "Scanning For Inbox Rules..."
                Start-Sleep -Seconds 2

                # Fetches Users With Possible Inbox Rules
                $selectUserMail = Get-InboxRule
                if ($selectUserMail.Count -eq 0){
                    Write-Host -ForegroundColor Red "There Weren't Any Inbox Rules To be Found."
                } else {
                    $selectUserMail | Out-GridView -Title "User Mailbox Rules" -PassThru
                    
                    foreach ($user in $selectUserMail){
                        Write-Host -ForegroundColor Yellow "Found User Inbox Rule For: $($user.displayName)`n"
                        Write-Host -ForegroundColor Yellow "Do You Want To Remove The Inbox Rule For $($user.displayName)? Y/N?"
                        $choice = Read-Host
                            if ($choice.ToUpper() -eq 'Y'){
                                Write-Host -ForegroundColor Yellow "Removing Inbox Rule For $($user.displayName).."
                                Get-InboxRule -Mailbox $user.Id | Disable-InboxRule -UserId $user.Id -WhatIf
                                Write-Host -ForegroundColor Green "Inbox Rule Removed."
                            }elseif($choice.ToUpper() -eq 'N'){
                                Write-Host -ForegroundColor Yellow "If No Other Users Found, Going Back To Home Menu.."
                            } else {
                                Write-Host -ForegroundColor Red "Error: Invalid Selection"
                            }
                    }
                }           
            }catch {
                Write-Host -ForegroundColor Red "Login Failed. Error: $($_.ExceptionMessage)"
                Write-Host -ForegroundColor Red "Detailed Error: $($_ | Out-String)"
            }
        Start-Sleep -Seconds 1
        Write-Host -ForegroundColor Cyan "Returning Back To Selection Menu."
        Start-Sleep -Seconds 1
        # Code Block for User License and Pulled Disabled Accounts
        }elseif ($userChoice -eq 2) {
            Connect-MgGraph -Scopes "User.Read.All" -NoWelcome -ErrorAction Stop

            $selectUsers = Get-MGUser -Filter 'accountEnabled eq false' -All | 
                        Select-Object displayName, Mail, Id, UserPrincipalName | ` 
                        Sort-Object -Property displayName | `
                        Out-GridView -Title "Disabled Account(s) With Possible Valid Licenses" -PassThru

            foreach ($user in $selectUsers) {
                $licenseDetail = Get-MgUserLicenseDetail -UserId $user.Id
                if ($licenseDetail.Count -eq 0) {
                    Write-Host -ForegroundColor Red "User: $($user.displayName) Does Not Have Any Valid Licenses."
                }else {
                    Write-Host -ForegroundColor Green "Found User License For: $($user.displayName)`nLicense Information Below:"
                    
                    $licenseDetail | Format-List

                    Write-Host -ForegroundColor Yellow "Would You Like To Remove $($user.displayName)'s License? Y/N?"
                    $choice = Read-Host
                        if ($choice.ToUpper() -eq 'Y' ){
                            Write-Host -ForegroundColor Yellow "To Remove License For $($user.displayName), Copy License Found Above For Its User."
                            Remove-MgUserLicenseDetail -UserId $user.Id -WhatIf
                        }elseif($choice.ToUpper() -eq 'N'){
                            Write-Host -ForegroundColor Yellow "Returning Other Users. If No Other Users Found, Returning Back To Home Menu.."
                        } else {
                            Write-Host -ForegroundColor Red "Error: Invalid Selection"
                        }
                }
            }
        Start-Sleep -Seconds 1
        Write-Host -ForegroundColor Cyan "Returning Back To Selection Menu."
        Start-Sleep -Seconds 1
        # Code Block for User Mailbox Conversion
        }elseif ($userChoice -eq 3){
            Connect-ExchangeOnline -UserPrincipalName upn_here  -Organization "domain.xyz" -ErrorAction Stop
            $selectUsersMailbox = Get-ExoMailbox | Select-Object displayName, PrimarySmtpAddress, RecipientTypeDetails | Sort-Object -Property RecipientTypeDetails | Out-GridView -Title "Select User Mailbox To Convert To Shared" -PassThru
        
            foreach ($user in $selectUsersMailbox){
                if ($user.RecipientTypeDetails -eq 'UserMailbox'){
                    Set-Mailbox -Identity $user.PrimarySmtpAddress -Type Shared
                    Write-Host -ForegroundColor Green "Converted $($user.displayName)'s Mailbox To Shared"
                } else{
                    Write-Host -ForegroundColor Yellow "$($user.displayName) Is Either Already A Shared Mailbox Or Is Not A Valid Mailbox."
                }
            }
        Start-Sleep -Seconds 1
        Write-Host -ForegroundColor Cyan "Returning Back To Selection Menu."
        Start-Sleep -Seconds 1
        # Code Block For Converting Back To Regular Mailbox
        }elseif ($userChoice -eq 4){
            Connect-ExchangeOnline -UserPrincipalName upn_here  -Organization "domain.xyz" -ErrorAction Stop
            $selectUsersMailbox = Get-ExoMailbox | Select-Object displayName, PrimarySmtpAddress, RecipientTypeDetails | Sort-Object -Property RecipientTypeDetails | Out-GridView -Title "Select User Mailbox To Convert Back To Regular" -PassThru
        
            foreach ($user in $selectUsersMailbox){
                if ($user.RecipientTypeDetails -eq 'SharedMailbox'){
                    Write-Host -ForegroundColor Yellow "Converting Shared Mailbox For $($user.displayName) Back To Regular Mailbox..."
                    try {
                        Set-Mailbox -Identity $user.PrimarySmtpAddress -Type Regular
                        Write-Host -ForegroundColor Green "$($user.displayName) Has Been Converted Back To A User Mailbox"
                    }
                    catch {
                        Write-Host -ForegroundColor Red "Failed To Convert $($user.displayName) Back To A User Mailbox. Error: $($_.ExceptionMessage)"
                    }
                } else{
                    Write-Host -ForegroundColor Cyan "$($user.displayName) Is Already A Regular Mailbox."
                }
            }
        Start-Sleep -Seconds 1
        Write-Host -ForegroundColor Cyan "Returning Back To Selection Menu."
        Start-Sleep -Seconds 1
        # Code Block To Revoke User Session 
        }elseif ($userChoice -eq 5){
            Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome -ErrorAction Stop

            $selectUserToRevokeSession = Get-MGUser -Filter 'accountEnabled eq false' -All | 
            Select-Object displayName, Mail, Id, UserPrincipalName | ` 
            Sort-Object -Property displayName | `
            Out-GridView -Title "Disabled Accounts To Revoke" -PassThru

            $selectUserToRevokeSession | ForEach-Object{
                Revoke-MgUserSignInSession -UserId $_.Id
                Write-Host -ForegroundColor Yellow "Revoked Session for: $($_.displayName) ($($_.UserPrincipalName))"
            }
        }
    } catch {
        Write-Host -ForegroundColor Red "$_ Enter A Valid Option. Try Again."
    }   
}