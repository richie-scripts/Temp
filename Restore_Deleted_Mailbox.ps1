<#
    Restore or Recreate a Mailbox

    Updated 14/07/2022
#>

# Users email address
$Mailbox = 'azhang@fmgl.com.au'




########
########
If (!$Mailbox) { $Mailbox = Read-Host 'Enter the mailbox to recover' }
Try { $null = New-Object System.Net.Mail.MailAddress($Mailbox) } Catch { Write-Warning "$Mailbox is not a valid email address."; return }

$OriginalPref = $ProgressPreference
$ProgressPreference = 'SilentlyContinue'
$DC = 'PTHADDS01.fmg.local'
Get-Variable -Name Inactive*, Active*, Recreate*, Restore* | Clear-Variable
$RestoreActiveMailbox = $false
$RestoreMailbox = $false

Connect-FMG365 -Reconnect
Connect-FMGExch -Reconnect
Connect-FMGMSOL -SilentAlreadyConnected
Connect-FMGTeams -SilentAlreadyConnected
Write-Host -ForegroundColor Green "Running restore on $Mailbox"

function Show-YesNoForm {
    [CmdletBinding()]
    [OutputType([Boolean])]
    param (
        [Parameter()][String]$Title = '',
        [Parameter()][String]$Message = ''
    )
    $null = [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $decision = [Microsoft.VisualBasic.Interaction]::MsgBox($Message, 'YesNo,SystemModal,Question', $Title)
    if ($decision -eq 'Yes') { return $true } else { return $false }
}

Function Show-MultiSelectForm {
    [CmdletBinding()]
    param (
        [Parameter()][String]$Title = '',
        [Parameter()][String]$Message = '',
        [Parameter()]$Options = ''
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = $Title
    $Form.ClientSize = '300,300'
    $Form.StartPosition = 'CenterScreen'
    $Form.FormBorderStyle = 'FixedDialog'
    $Form.MaximizeBox = $false
    $Form.MinimizeBox = $false
    $Form.Topmost = $True

    $Label = New-Object System.Windows.Forms.Label
    $Label.Location = New-Object System.Drawing.Point(10, 10)
    $Label.Size = New-Object System.Drawing.Size(($Form.ClientSize.Width - 20), ($Message.Split([Environment]::NewLine).Count * 12))
    $Label.Text = $Message

    $ListBox = New-Object System.Windows.Forms.Listbox
    $ListBox.Location = New-Object System.Drawing.Point(10, ($Label.Location.Y + $Label.Size.Height + 5))
    $ListBox.Size = New-Object System.Drawing.Size(($Form.ClientSize.Width - 20), ((20 * $Options.Count) + 10))
    $ListBox.SelectionMode = 'MultiExtended'
    $ListBox.Font = New-Object System.Drawing.Font('Consolas', 8, [System.Drawing.FontStyle]::Regular)
    ForEach ($Option in $Options) {
        [void]$ListBox.Items.Add("$Option")
    }

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $CancelButton.Location = New-Object System.Drawing.Point(($Form.ClientSize.Width - $CancelButton.Size.Width - 10), ($ListBox.Location.Y + $ListBox.Size.Height + 5))
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Size = New-Object System.Drawing.Size(75, 23)
    $OKButton.Location = New-Object System.Drawing.Point(($CancelButton.Location.X - $CancelButton.Size.Width - 5), ($ListBox.Location.Y + $ListBox.Size.Height + 5))
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $Form.AcceptButton = $OKButton
    $Form.CancelButton = $CancelButton
    $Form.Controls.AddRange(@($Label, $ListBox, $OKButton, $CancelButton))
    $Form.ClientSize = "$($Form.ClientSize.Width),$($OKButton.Location.Y + $OKButton.Size.Height + 10)"

    $Result = $Form.ShowDialog()
    If ($Result -eq [System.Windows.Forms.DialogResult]::OK) { return $ListBox.SelectedItems }
    Else { return $null }
}

# Check if currently active mailbox exists and warn about deleting it
Write-Host -ForegroundColor Green "`nChecking for an active mailbox..."
$ActiveMailbox = Get-EXOMailbox -Identity $Mailbox -PropertySets StatisticsSeed, Archive, SoftDelete, Minimum -ErrorAction SilentlyContinue
If ($ActiveMailbox) {
    $ActiveMailboxStats = $ActiveMailbox | Get-EXOMailboxStatistics
    Write-Host -ForegroundColor Cyan '- Found an active mailbox! Do you you want to recreate this mailbox or just restore old emails?'
    $RecreateActiveMailbox = Show-YesNoForm 'Delete active mailbox?' "There is an active mailbox for $Mailbox `nDo you want to delete it? `n`nMailbox size: $(($ActiveMailboxStats.TotalItemSize -Split '\('))[0]) `nItem count:   $($ActiveMailboxStats.ItemCount) `n`nYes = Delete and recreate users mailbox `nNo = Move on to restoring emails"
    Write-Host -ForegroundColor Cyan "-- You said $(If ($RecreateActiveMailbox){'yes'}else{'no'}) to deleting the active mailbox! $(Switch(1..4|Get-Random){1{'Brave'}2{'Questionable'}3{'Interesting'}4{'Juicy'}}) choice!"
}
Else { Write-Host -ForegroundColor Cyan '- None found. Moving on...'; $RecreateActiveMailbox = $true }

If ($RecreateActiveMailbox) {
    Write-Host -ForegroundColor Green "`n$(If($ActiveMailbox){'Recreating'}else{'Creating'}) user mailbox..."

    # Remove AD values
    $ADUser = Get-ADUser -Filter { (EmailAddress -eq $Mailbox) -or (UserPrincipalName -eq $Mailbox) } -Server $DC -Properties *
    If ($ADUser) {
        Write-Host -ForegroundColor Cyan '- Removing AD values'
        $ADUser.GetEnumerator() | Where-Object { $_.Key -like 'msExch*' } | Select-Object -ExpandProperty Key | ForEach-Object { Set-ADUser -Identity $ADUser -Server $DC -ErrorAction SilentlyContinue -Clear $_ }
        Start-Sleep 10
    }

    # Get users Teams groups to add back later
    Try {
        Write-Host -ForegroundColor Cyan '- Getting list of current Teams access to re-add'
        Connect-FMGTeams -Silent
        $TeamsGroups = @()
        ForEach ($Team in (Get-Team -User $Mailbox -ErrorAction Stop)) {
            $thisTeamPerms = $Team | Get-TeamUser | Where-Object { $_.User -eq $Mailbox }
            $TeamsGroups += [PSCustomObject][ordered]@{
                GroupId     = $Team.GroupId
                DisplayName = $Team.DisplayName
                User        = $thisTeamPerms.User
                Role        = $thisTeamPerms.Role
            }
        }
    }
    Catch {}
    Start-Sleep 10

    # Remove user from MSOL
    Try {
        Get-MsolUser -UserPrincipalName $Mailbox -ErrorAction Stop > $null
        Write-Host -ForegroundColor Cyan '- Removing user from MSOL'
        Remove-MsolUser -UserPrincipalName $Mailbox -Force > $null
        Start-Sleep 10
        Remove-MsolUser -UserPrincipalName $Mailbox -RemoveFromRecycleBin -Force > $null
    }
    Catch {}

    # Create Mailbox On Prem if it doesn't exist and enable the archive
    Connect-FMGExch -Silent
    Try { Get-RemoteMailbox $Mailbox -DomainController $DC -ErrorAction Stop > $null }
    Catch {
        Write-Host -ForegroundColor Cyan '- Enabling RemoteMailbox in on-prem Exchange'
        Enable-RemoteMailbox $($Mailbox.Split('@')[0]) -RemoteRoutingAddress "$($Mailbox.Split('@')[0])@service.fmgl.com.au" -PrimarySMTPAddress $Mailbox -DomainController $DC > $null
        Set-RemoteMailbox $($Mailbox.Split('@')[0]) -EmailAddressPolicyEnabled $True -DomainController $DC > $null
    }
    If ((Get-RemoteMailbox $Mailbox -DomainController $DC).ArchiveGuid -eq [System.Guid]::empty) {
        Write-Host -ForegroundColor Cyan '- Enabling archive for new mailbox'
        Enable-RemoteMailbox $Mailbox -DomainController $DC -Archive -ErrorAction SilentlyContinue > $null
    }

    If ($ActiveMailbox) { $RestoreActiveMailbox = $true }
}

# Check if mailbox has been soft deleted to recover
Write-Host -ForegroundColor Green "`nChecking for soft deleted mailboxes..."
$InactiveMailbox = Get-365Mailbox -SoftDeletedMailbox -Identity $Mailbox -ErrorAction SilentlyContinue | Sort-Object -Descending WhenSoftDeleted
If ($InactiveMailbox) {
    $InactiveMailboxStats = $InactiveMailbox | ForEach-Object { Get-365MailboxStatistics -IncludeSoftDeletedRecipients $_.ExchangeGuid.toString() }
    Write-Host -ForegroundColor Cyan "- Found $($InactiveMailbox.Name.Count) soft deleted mailboxes! Select any you want to restore."
    $Message = "Found $($InactiveMailbox.Name.Count) soft deleted mailbox$(If($InactiveMailbox.Name.Count -gt 1){'es'})!`nSelect any you want to restore.`n `nDate Deleted       Inbox         Deleted Item          ID `n                              Size                Size"
    $Options = ForEach ($Item in $InactiveMailbox) {
        $Stats = $InactiveMailboxStats | Where-Object { $_.MailboxGuid -eq $Item.ExchangeGuid }
        $ExchID = $Item.ExchangeGuid.ToString().Substring(0, $Item.ExchangeGuid.ToString().IndexOf('-'))
        $InboxSize = $Stats.TotalItemSize.Value.ToString().Substring(0, ($Stats.TotalItemSize.Value.ToString().IndexOf('(') - 1))
        $DeletedItemSize = $Stats.TotalDeletedItemSize.Value.ToString().Substring(0, ($Stats.TotalDeletedItemSize.Value.ToString().IndexOf('(') - 1))
        "$($Item.WhenSoftDeleted.ToString('dd/MM/yyyy')) | $($InboxSize)$(' '*(9-$InboxSize.Length))| $($DeletedItemSize)$(' '*(9-$DeletedItemSize.Length))| $($ExchID)"
    }
    $RestoreItems = Show-MultiSelectForm -Title $Mailbox -Message $Message -Options $Options
    $RestoreMailbox = @()
    $RestoreItemsCount = 0
    ForEach ($Item in $RestoreItems) {
        $RestoreItemsCount++
        $ItemID = $Item.Split('|')[-1].Trim()
        $RestoreMailbox += $InactiveMailbox | Where-Object { $_.ExchangeGuid.ToString().Substring(0, $_.ExchangeGuid.ToString().IndexOf('-')) -eq $ItemID }
    }
    Write-Host -ForegroundColor Cyan "-- You selected $(If ($RestoreItemsCount -ge 1){$RestoreItemsCount}else{'a big fat zero'}) mailbox$(If($RestoreItemsCount -gt 1 -or $RestoreItemsCount -eq 0){'es'}) to restore! $(Switch(1..4|Get-Random){1{'Wonderful'}2{'Solid'}3{'Great'}4{'Spectacular'}}) choice!"
}
Else { Write-Host -ForegroundColor Cyan '- No inactive mailboxes were found to restore.' }

If ($RestoreMailbox -or $RestoreActiveMailbox -or $TeamsGroups) {
    # Check new mailbox exists and has a new Guid
    If ($RecreateActiveMailbox) {
        Write-Host -ForegroundColor Cyan -NoNewline "`nWaiting for next AzureAD Sync. This can take up to 20 mins.`nTo force a sync, RDP to IOCOPSSYNC01.fmg.local and run C:\Scripts\RunAzureADSync.ps1`nSearching for new mailbox..."
    }
    Do {
        If ($RecreateActiveMailbox) {
            Write-Host -ForegroundColor Cyan -NoNewline '.'
            Start-Sleep -s 15
        }
        Start-Sleep 5
        $NewMailbox = Get-EXOMailbox $Mailbox -PropertySets StatisticsSeed, Archive, Minimum -ErrorAction SilentlyContinue
        $NewMailbox365 = Get-365Mailbox $Mailbox -ErrorAction SilentlyContinue
    }
    Until((($RecreateActiveMailbox) -and ($NewMailbox) -and ($NewMailbox.ExchangeGuid -ne [System.Guid]::empty) -and ($NewMailbox.ExchangeGuid -ne $ActiveMailbox.ExchangeGuid) -and ($NewMailbox.ArchiveGuid -ne [System.Guid]::empty) -and ($NewMailbox.ArchiveGuid -ne $ActiveMailbox.ArchiveGuid) -and ($NewMailbox.ExchangeGuid -eq $NewMailbox365.ExchangeGuid)) -or ((!$RecreateActiveMailbox) -and ($NewMailbox)))
    If ($RecreateActiveMailbox) {
        Start-Sleep 60
        Write-Host -ForegroundColor Cyan '. Found.'
    }

    # Restore Teams access
    If (($TeamsGroups | Measure-Object).Count -gt 0) {
        Write-Host -ForegroundColor Green 'Restoring MS Teams access'
        Connect-FMGTeams -Silent
        ForEach ($Team in $TeamsGroups) {
            Try {
                Add-TeamUser -GroupId $Team.GroupId -User $Team.User -Role $Team.Role -EA Stop
                Write-Host -ForegroundColor Cyan "- Added to team `'$($Team.DisplayName)`' as role `'$($Team.Role)`'"
            }
            Catch {
                Write-Host -ForegroundColor Red "- Unable to add to team `'$($Team.DisplayName)`' as role `'$($Team.Role)`'`n-- $($_.Exception.Message)"
            }
        }
    }

    # Restore the old mailbox
    Connect-FMG365 -SilentAlreadyConnected

    If ($RestoreMailbox) {
        Write-Host -ForegroundColor Green 'Restoring inactive mailboxes'
        $RestoreMailbox | ForEach-Object {
            Try {
                $RestoreTime = (Get-Date).tostring('dd.MM.yy-HH.mm.ss')
                New-365MailboxRestoreRequest -Name "MailRestore-$RestoreTime-$Mailbox" -SourceMailbox $($_.ExchangeGuid.toString()) -TargetMailbox $($NewMailbox.ExchangeGuid.toString()) -AllowLegacyDNMismatch -ConflictResolutionOption KeepSourceItem -BadItemLimit 5000 -AcceptLargeDataLoss -WarningAction SilentlyContinue -ErrorAction Stop > $null
                Write-Host -ForegroundColor Cyan "- Created mailbox restore request `"MailRestore-$RestoreTime-$Mailbox`""
            }
            Catch { Write-Host -ForegroundColor Red "- Unable to create mailbox restore request`n$($_.Exception.Message)" }
            Start-Sleep 1
        }

        If (($RestoreMailbox | Where-Object { $_.ArchiveGuid -ne [System.Guid]::empty })) {
            Write-Host -ForegroundColor Green 'Restoring inactive archives'
            $RestoreMailbox | Where-Object { $_.ArchiveGuid -ne [System.Guid]::empty } | ForEach-Object {
                Try {
                    $RestoreTime = (Get-Date).tostring('dd.MM.yy-HH.mm.ss')
                    New-365MailboxRestoreRequest -Name "ArchiveRestore-$RestoreTime-$Mailbox" -SourceMailbox $($_.ExchangeGuid.toString()) -SourceIsArchive -TargetMailbox $($NewMailbox.ExchangeGuid.toString()) -TargetIsArchive -TargetRootFolder 'Archive' -AllowLegacyDNMismatch -ConflictResolutionOption KeepSourceItem -BadItemLimit 5000 -AcceptLargeDataLoss -WarningAction SilentlyContinue -ErrorAction Stop > $null
                    Write-Host -ForegroundColor Cyan "- Created archive restore request `"ArchiveRestore-$RestoreTime-$Mailbox`""
                }
                Catch { Write-Host -ForegroundColor Red "- Unable to create archive restore request`n$($_.Exception.Message)" }
            }
        }
    }

    If ($RestoreActiveMailbox) {
        $OldActiveMailbox = (Get-365Mailbox -Identity $Mailbox -SoftDeletedMailbox | Sort-Object -Descending WhenSoftDeleted)[0]
        Write-Host -ForegroundColor Green 'Restoring recently active mailbox'

        Try {
            $RestoreTime = (Get-Date).tostring('dd.MM.yy-HH.mm.ss')
            New-365MailboxRestoreRequest -Name "MailRestore-$RestoreTime-$Mailbox" -SourceMailbox $($OldActiveMailbox.ExchangeGuid.toString()) -TargetMailbox $($NewMailbox.ExchangeGuid.toString()) -AllowLegacyDNMismatch -ConflictResolutionOption KeepSourceItem -BadItemLimit 5000 -AcceptLargeDataLoss -WarningAction SilentlyContinue -ErrorAction Stop > $null
            Write-Host -ForegroundColor Cyan "- Created mailbox restore request `"MailRestore-$RestoreTime-$Mailbox`""
        }
        Catch { Write-Host -ForegroundColor Red "- Unable to create mailbox restore request`n$($_.Exception.Message)" }
        Start-Sleep 1

        If ($OldActiveMailbox.ArchiveGuid -ne [System.Guid]::empty) {
            Write-Host -ForegroundColor Green 'Restoring archives'
            Try {
                $RestoreTime = (Get-Date).tostring('dd.MM.yy-HH.mm.ss')
                New-365MailboxRestoreRequest -Name "ArchiveRestore-$RestoreTime-$Mailbox" -SourceMailbox $($OldActiveMailbox.ExchangeGuid.toString()) -SourceIsArchive -TargetMailbox $($NewMailbox.ExchangeGuid.toString()) -TargetIsArchive -TargetRootFolder 'Archive' -AllowLegacyDNMismatch -ConflictResolutionOption KeepSourceItem -BadItemLimit 5000 -AcceptLargeDataLoss -WarningAction SilentlyContinue -ErrorAction Stop > $null
                Write-Host -ForegroundColor Cyan "- Created archive restore request `"ArchiveRestore-$RestoreTime-$Mailbox`""
            }
            Catch { Write-Host -ForegroundColor Red "- Unable to create archive restore request`n$($_.Exception.Message)" }
        }
    }

    If ($RestoreMailbox -or $RestoreActiveMailbox) {
        Write-Host -ForegroundColor Green "`nTo check the status of the restore processes, use:"
        Write-Host -ForegroundColor Cyan "- Get-365MailboxRestoreRequest -TargetMailbox `"$($NewMailbox.ExchangeGuid.toString())`"`n"
        Write-Host -ForegroundColor Green 'If you are getting restore failures, use this to check the errors:'
        Write-Host -ForegroundColor Cyan "- Get-365MailboxRestoreRequest -TargetMailbox `"$($NewMailbox.ExchangeGuid.toString())`" -Status Failed | Select -Last 1 -ExpandProperty RequestGuid | Get-365MailboxRestoreRequestStatistics -IncludeReport | Select -ExpandProperty LastFailure"
        Write-Host -ForegroundColor Green 'Then you can resume the restore request:'
        Write-Host -ForegroundColor Cyan "- Get-365MailboxRestoreRequest -TargetMailbox `"$($NewMailbox.ExchangeGuid.toString())`" -Status Failed | Select -Last 1 -ExpandProperty RequestGuid | Resume-365MailboxRestoreRequest"
    }
}
Else {
    Write-Host -ForegroundColor Green  "`nMailbox will be ready on next AzureAD sync. This takes about 5-20 mins.`nSpeed this up by manually running a sync on IOCOPSSYNC01.fmg.local > C:\Scripts\RunAzureADSync.ps1"
}
Write-Host -ForegroundColor Green "`nTo paste into Incident:"
Write-Host -ForegroundColor Cyan "Mailbox $Mailbox has been restored.`nThe mailbox should now be accessible. If they are still getting Outlook errors or bad user password errors, recreate the local computer profile and reset their password."

$ProgressPreference = $OriginalPref
