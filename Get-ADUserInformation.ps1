<#
.SYNOPSIS
    The script get all information about the users in AD.
.DESCRIPTION
    This script uses the Get-ADUser and  Get-ADPrincipalGroupMembership cmdlet to search information about the users.

    Conventions:
    - Active Directory (AD)

    Menu for automatic guide for Get-ADUserReport:
    [0] - Report AD Users

.NOTES
    History:
    Version     Who                 When            What
    1.0         Julian Gómez        02/15/2024      - Creation of script.
    
.EXAMPLE
    Get-ADUserReport
    Only run the script, it will guide you through the process.
#>

function Get-ADUserReport {

    # Declare variables
    $UserPath = "C:\Users\$env:UserName\Documents"      # Path User Documents
    $Date = Get-Date -Format "MMddyyyy_HHmm"            # Current Date
    $LoginUser = $env:UserName                          # Current User
    $l = "=" * 120                                      # Format Separator
    $s = "    "                                         # Format Space
    $Id = 0                                             # Counter
    
    #region // ================================================================================================ [ Function - Set Global Variables ]
    function Initialize-GlobalVariablesAD {
        $Global:ReportADUser = @()                      # Array Report Users
    }

    #region // ================================================================================================ [ Check Module - Excel ]
    Function Install-ModuleExcel {
        if ($null -eq (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host $s"Excel module is requied, do you want to install it?" -F Yellow
                    
            $install = Read-Host Do you want to install module? [Y] Yes [N] No 
            if ($install -match "[yY]") { 
                Write-Host $s"Installing Import-Excel module" -F Cyan
                Install-Module -Name ImportExcel -Force
            }
            else {
                Write-Error "Please install Import-Excel module."
            }
        }
    }
    
    #region // ================================================ [ Function - [Back Menu] ]
    function Wait-Redirection($Menu, [int]$Timer) {
        for ($sec = $Timer; $sec -gt -1; $sec--) {
            Write-Host -NoNewLine ("`r$s Back to $Menu..." + (": {0:d2}" -f $sec)) -F DarkYellow
            Start-Sleep -Seconds 1
        }
        Write-Host "`r" -NoNewLine
        if ($Menu -eq "Menu") {
            Approve-MenuADUser
        }
        else {
            Approve-MenuADUser
        }
       
    }
    #region // ================================================================================================ [ Function - Get AdUsers ]
    function Get-ADUserInformation {
        Write-Host $s"Generated the report. Wait!." -F DarkGray

        Try { 
            $ResultSetADUsers = Get-ADUser -Filter * -ResultSetSize 100 -EA Stop
        }
        catch [System.Exception] {
            if ($_.Exception -match "No connection available") {
                Write-Host $s"Was unable to connect to Exchange Online" -F Cyan
            }
            else {
                Write-Host $s""$_.Exception
            }
        }
    
        $CountResultSet = $ResultSetADUsers.Count
        $CountUser = 1
        foreach ($vUser in $ResultSetADUsers) {
                Write-Host -NoNewLine ("`r$s Progress... [" + ($CountUser -f "") + " / $CountResultSet]") -F DarkYellow
                Try { 
                    $ResultSetGroupsMemmbership = Get-ADPrincipalGroupMembership $vUser | Select-Object -ExpandProperty Name
                }
                catch [System.Exception] {
                    if ($_.Exception -match "No connection available") {
                        Write-Host $s"Was unable to connect to Exchange Online" -F Cyan
                    }
                    else {
                        Write-Host $s""$_.Exception
                    }
                }
        
                $Global:ReportADUser += [PSCustomObject]@{
                    SamAccountName    = $vUser.sAMAccountName
                    Name              = $vUser.Name
                    UserPrincipalName = $vUser.userPrincipalName
                    Status            = $vUser.Enabled
                    OU                = $vUser.distinguishedName
                    MemberOf          = $ResultSetGroupsMemmbership | Out-String -Verbose
                }
                $CountUser++
        }

        $Global:ReportADUser | Export-Excel -Path "$UserPath\ReportADUsers.xlsx" -WorksheetName "ADUsers" -AutoSize -TableName "TD_ADUsers"  -TableStyle "Medium6" -BoldTopRow
        Write-Host "|| " + $l -F DarkGreen
        Write-Host $s"A report of the AD User information has been created!." -F Green
        Write-Host $s"Report path: [$UserPath\ReportADUsers.xlsx]." -F Green
        Write-Host "|| " + $l -F DarkGreen
        Wait-Redirection "Menu" 2
    }

    #region // ================================================ [ Active Directory - [Menu]]
    function Approve-MenuADUser {
        Write-Host "//" + $l "[ Menu Active Directory ]" -F DarkGreen
        Write-Host $s"In this section you will be able to make some modifications to the user." -F Gray
    
        Write-Host $s$l -F Magenta
        $(Write-Host $s"[0] - " -F White -NoNewline) + $(Write-Host "Report AD Users" -F White)
        # $(Write-Host $s"[0] - " -F White -NoNewline) + $(Write-Host "[SubMenu] - " -F Magenta -NoNewline) + $(Write-Host "Direct Reports" -F White)
        Write-Host $s$l -F Magenta
        Write-Host $s
        $op_Menu = $(Write-Host $s"Exit [any key] | Type Option (#): " -F Cyan -NoNewline; Read-Host)
    
        switch ($op_Menu) {
            0 { Get-ADUserInformation }
            Default {
                Clear-Host
                Write-Host "//" + $l "[ Finish <<< ]" -F DarkGreen
                Initialize-GlobalVariablesAD
                Exit
            }
        }
    }

    #region // ================================================ [ Active Directory - [Launch Initial Code] ]
    Write-Host "//" + $l "[ Start >>> ]" -F DarkGreen
    Write-Host $s"Type Request: (AD) Active Directory." -F DarkGray
    Initialize-GlobalVariablesAD
    Install-ModuleExcel
    Approve-MenuADUser
}
Get-ADUserReport
