 Function Start-PmWUInstall
{
    [CmdletBinding(
        SupportsShouldProcess=$True,
        ConfirmImpact="High"
    )]    
    Param
    (
        #Pre search criteria
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [ValidateSet("Driver", "Software")]
        [String]$UpdateType="",
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [String[]]$UpdateID,
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [Int]$RevisionNumber,
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [String[]]$CategoryIDs,
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [Switch]$IsInstalled,
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [Switch]$IsHidden,
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [Switch]$WithHidden,
        [String]$Criteria,
        [Switch]$ShowSearchCriteria,
        
        #Post search criteria
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [String[]]$Category="",
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [String[]]$KBArticleID,
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [String]$Title,
        
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [String[]]$NotCategory="",
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [String[]]$NotKBArticleID,
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [String]$NotTitle,
        
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [Alias("Silent")]
        [Switch]$IgnoreUserInput,
        [parameter(ValueFromPipelineByPropertyName=$true)]
        [Switch]$IgnoreRebootRequired,
        
        #Connection options
        [String]$ServiceID,
        [Switch]$WindowsUpdate,
        [Switch]$MicrosoftUpdate,
        
        #Mode options
        [Switch]$ListOnly,
        [Switch]$DownloadOnly,
        [Alias("All")]
        [Switch]$AcceptAll,
        [Switch]$AutoReboot,
        [Switch]$IgnoreReboot,
        [Switch]$AutoSelectOnly,
        [Switch]$Debuger
        
    )

    Begin
    {
        $identity        = [Security.Principal.WindowsIdentity]::GetCurrent()
        $currentUserName = $identity.Name

        $User = [Security.Principal.WindowsIdentity]::GetCurrent()
        $Role = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

        if(!$Role)
        {
            if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
            {
                Write-PmLog -Level Information -Message "To perform some operations you must run an elevated Windows PowerShell console."
            }
            else
            {
                Write-PmLog -Level Information -Message "To perform some operations you must run an elevated Windows PowerShell console."
            }
        } #End If !$Role
    }

    Process
    {
        <#
            Start STAGE 0: Prepare environment 
        #>

        if ($Debuger) { Write-PmLog -Level Debug -Message "STAGE 0: Prepare environment" }
        If($IsInstalled)
        {
            $ListOnly = $true
            if ($Debuger) { Write-PmLog -Level Debug -Message "Change to ListOnly mode" }
        } #End If $IsInstalled

        if ($Debuger) { Write-PmLog -Level Debug -Message "Check reboot status only for local instance" }
        Try
        {
            $objSystemInfo = New-Object -ComObject "Microsoft.Update.SystemInfo"    
            If($objSystemInfo.RebootRequired)
            {
                if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                {
                    Write-PmLog -Level Information -Message "Reboot is required to continue"
                }
                else
                {
                    Write-PmLog -Level Information -Message "Reboot is required to continue"
                }

                If($AutoReboot)
                {
                    Restart-ProComputer -TimeOut 300
                } #End If $AutoReboot

                If(!$ListOnly)
                {
                    Return
                } #End If !$ListOnly    
                
            } #End If $objSystemInfo.RebootRequired
        } #End Try
        Catch
        {
            if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
            {
                Write-PmLog -Level Information -Message "Support local instance only, Continue..."
            }
            else
            {
                Write-PmLog -Level Information -Message "Support local instance only, Continue..."
            }
        } #End Catch
        
        if ($Debuger) { Write-PmLog -Level Debug -Message "Set number of stage" }
        If($ListOnly)
        {
            $NumberOfStage = 1
        } 
        ElseIf($DownloadOnly)
        {
            $NumberOfStage = 3
        } 
        Else
        {
            $NumberOfStage = 4
        } 
        
        <#           
         End STAGE 0: Prepare environment 
        #>
        
        <#
          Start STAGE 1: Get updates list 
        #>     
        $sMessage = '[1/' + $NumberOfStage + '] Get available updates...'
        if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
        {
            Write-PmLog -Level Information -Message $sMessage
        }
        else
        {
            Write-PmLog -Level Information -Message $sMessage
        }

        if ($Debuger) { Write-PmLog -Level Debug -Message "STAGE 1: Get updates list" }
        if ($Debuger) { Write-PmLog -Level Debug -Message "Create Microsoft.Update.ServiceManager object" }
        $objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" 
        
        if ($Debuger) { Write-PmLog -Level Debug -Message "Create Microsoft.Update.Session object" }
        $objSession = New-Object -ComObject "Microsoft.Update.Session" 
        
        if ($Debuger) { Write-PmLog -Level Debug -Message "Create Microsoft.Update.Session.Searcher object" }
        $objSearcher = $objSession.CreateUpdateSearcher()

        If($WindowsUpdate)
        {
            if ($Debuger) { Write-PmLog -Level Debug -Message "Set source of updates to Windows Update" }
            $objSearcher.ServerSelection = 2
            $serviceName = "Windows Update"
        } #End If $WindowsUpdate
        ElseIf($MicrosoftUpdate)
        {
            if ($Debuger) { Write-PmLog -Level Debug -Message "Set source of updates to Microsoft Update" }
            $serviceName = $null
            Foreach ($objService in $objServiceManager.Services) 
            {
                If($objService.Name -eq "Microsoft Update")
                {
                    $objSearcher.ServerSelection = 3
                    $objSearcher.ServiceID = $objService.ServiceID
                    $serviceName = $objService.Name
                    Break
                }#End If $objService.Name -eq "Microsoft Update"
            }#End ForEach $objService in $objServiceManager.Services
            
            If(-not $serviceName)
            {
                if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                {
                    Write-PmLog -Level Information -Message "Can't find registered service Microsoft Update. Use Get-WUServiceManager to get registered service."
                }
                else
                {
                    Write-PmLog -Level Information -Message "Can't find registered service Microsoft Update. Use Get-WUServiceManager to get registered service."
                }
                
                Return
            }#Enf If -not $serviceName
        } #End Else $WindowsUpdate If $MicrosoftUpdate
        Else
        {
            Foreach ($objService in $objServiceManager.Services) 
            {
                If($ServiceID)
                {
                    If($objService.ServiceID -eq $ServiceID)
                    {
                        $objSearcher.ServiceID = $ServiceID
                        $objSearcher.ServerSelection = 3
                        $serviceName = $objService.Name
                        Break
                    } #End If $objService.ServiceID -eq $ServiceID
                } #End If $ServiceID
                Else
                {
                    If($objService.IsDefaultAUService -eq $True)
                    {
                        $serviceName = $objService.Name
                        Break
                    } #End If $objService.IsDefaultAUService -eq $True
                } #End Else $ServiceID
            } #End Foreach $objService in $objServiceManager.Services
        } #End Else $MicrosoftUpdate

        if ($Debuger) { Write-PmLog -Level Debug -Message "Set source of updates to $serviceName" }
        if ($Debuger) { Write-PmLog -Level Debug -Message "Connecting to $serviceName server. Please wait..." }

        Try
        {
            $search = ""
            
            If($Criteria)
            {
                $search = $Criteria
            } #End If $Criteria
            Else
            {
                If($IsInstalled) 
                {
                    $search = "IsInstalled = 1"
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: IsInstalled = 1" }
                } #End If $IsInstalled
                Else
                {
                    $search = "IsInstalled = 0"    
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: IsInstalled = 0" }
                } #End Else $IsInstalled
                
                If($UpdateType -ne "")
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: Type = $UpdateType" }
                    $search += " and Type = '$UpdateType'"
                } #End If $UpdateType -ne ""                    
                
                If($UpdateID)
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: UpdateID = '$([string]::join(", ", $UpdateID))'" }
                    $tmp = $search
                    $search = ""
                    $LoopCount = 0
                    Foreach($ID in $UpdateID)
                    {
                        If($LoopCount -gt 0)
                        {
                            $search += " or "
                        } #End If $LoopCount -gt 0
                        If($RevisionNumber)
                        {
                            if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: RevisionNumber = '$RevisionNumber'"   }  
                            $search += "($tmp and UpdateID = '$ID' and RevisionNumber = $RevisionNumber)"
                        } #End If $RevisionNumber
                        Else
                        {
                            $search += "($tmp and UpdateID = '$ID')"
                        } #End Else $RevisionNumber
                        $LoopCount++
                    } #End Foreach $ID in $UpdateID
                } #End If $UpdateID

                If($CategoryIDs)
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: CategoryIDs = '$([string]::join(", ", $CategoryIDs))'" }
                    $tmp = $search
                    $search = ""
                    $LoopCount =0
                    Foreach($ID in $CategoryIDs)
                    {
                        If($LoopCount -gt 0)
                        {
                            $search += " or "
                        } #End If $LoopCount -gt 0
                        $search += "($tmp and CategoryIDs contains '$ID')"
                        $LoopCount++
                    } #End Foreach $ID in $CategoryIDs
                } #End If $CategoryIDs
                
                If($IsHidden) 
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: IsHidden = 1" }
                    $search += " and IsHidden = 1"    
                } #End If $IsNotHidden
                ElseIf($WithHidden) 
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: IsHidden = 1 and IsHidden = 0" }
                } #End ElseIf $WithHidden
                Else
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: IsHidden = 0" }
                    $search += " and IsHidden = 0"    
                } #End Else $WithHidden
                
                #Don't know why every update have RebootRequired=false which is not always true
                If($IgnoreRebootRequired) 
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Set pre search criteria: RebootRequired = 0" }
                    $search += " and RebootRequired = 0"    
                } #End If $IgnoreRebootRequired
            } #End Else $Criteria
            
            
            if ($Debuger) { Write-PmLog -Level Debug -Message "Search criteria is: $search" }

            If($ShowSearchCriteria)
            {
            if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
            {
                Write-PmLog -Level Information -Message "Search criteria is: $search"
            }
            else
            {
                Write-PmLog -Level Information -Message "Search criteria is: $search"
            }
                
            } 
            
            $objResults = $objSearcher.Search($search)
        } #End Try
        Catch
        {
            If($_ -match "HRESULT: 0x80072EE2")
            {
                if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                {
                    Write-PmLog -Level Information -Message "Probably you don't have connection to Windows Update server"
                }
                else
                {
                    Write-PmLog -Level Information -Message "Probably you don't have connection to Windows Update server"
                }
                
            } #End If $_ -match "HRESULT: 0x80072EE2"
            Return
        } #End Catch

        $objCollectionUpdate = New-Object -ComObject "Microsoft.Update.UpdateColl" 
        
        $NumberOfUpdate = 1
        $UpdateCollection = @()
        $UpdatesExtraDataCollection = @{}
        $PreFoundUpdatesToDownload = $objResults.Updates.count

        if ($Debuger) { Write-PmLog -Level Debug -Message "Found [$PreFoundUpdatesToDownload] Updates in pre search criteria"   }              

        Foreach($Update in $objResults.Updates)
        {    
            $UpdateAccess = $true
  
            if ($Debuger) { Write-PmLog -Level Debug -Message "Set post search criteria: $($Update.Title)" }
            
            If($Category -ne "")
            {
                $UpdateCategories = $Update.Categories | Select-Object Name
                if ($Debuger) { Write-PmLog -Level Debug -Message "Set post search criteria: Categories = '$([string]::join(", ", $Category))'"  }   
                Foreach($Cat in $Category)
                {
                    If(!($UpdateCategories -match $Cat))
                    {
                        if ($Debuger) { Write-PmLog -Level Debug -Message "UpdateAccess: false" }
                        $UpdateAccess = $false
                    } #End If !($UpdateCategories -match $Cat)
                    Else
                    {
                        $UpdateAccess = $true
                        Break
                    } #End Else !($UpdateCategories -match $Cat)
                } #End Foreach $Cat in $Category    
            } #End If $Category -ne ""

            If($NotCategory -ne "" -and $UpdateAccess -eq $true)
            {
                $UpdateCategories = $Update.Categories | Select-Object Name
                if ($Debuger) { Write-PmLog -Level Debug -Message "Set post search criteria: NotCategories = '$([string]::join(", ", $NotCategory))'"  }  
                Foreach($Cat in $NotCategory)
                {
                    If($UpdateCategories -match $Cat)
                    {
                        if ($Debuger) { Write-PmLog -Level Debug -Message "UpdateAccess: false" }
                        $UpdateAccess = $false
                        Break
                    } #End If $UpdateCategories -match $Cat
                } #End Foreach $Cat in $NotCategory    
            } #End If $NotCategory -ne "" -and $UpdateAccess -eq $true                    
            
            If($KBArticleID -ne $null -and $UpdateAccess -eq $true)
            {
                if ($Debuger) { Write-PmLog -Level Debug -Message "Set post search criteria: KBArticleIDs = '$([string]::join(", ", $KBArticleID))'" }
                If(!($KBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs))
                {
                    Write-PmLog -Level Debug -Message "UpdateAccess: false"
                    $UpdateAccess = $false
                } #End If !($KBArticleID -match $Update.KBArticleIDs)                                
            } #End If $KBArticleID -ne $null -and $UpdateAccess -eq $true

            If($NotKBArticleID -ne $null -and $UpdateAccess -eq $true)
            {
                if ($Debuger) { Write-PmLog -Level Debug -Message "Set post search criteria: NotKBArticleIDs = '$([string]::join(", ", $NotKBArticleID))'" }
                If($NotKBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs)
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "UpdateAccess: false" }
                    $UpdateAccess = $false
                } #End If$NotKBArticleID -match $Update.KBArticleIDs -and "" -ne $Update.KBArticleIDs                    
            } #End If $NotKBArticleID -ne $null -and $UpdateAccess -eq $true
            
            If($Title -and $UpdateAccess -eq $true)
            {
                if ($Debuger) { Write-PmLog -Level Debug -Message "Set post search criteria: Title = '$Title'" }
                If($Update.Title -notmatch $Title)
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "UpdateAccess: false" }
                    $UpdateAccess = $false
                } #End If $Update.Title -notmatch $Title
            } #End If $Title -and $UpdateAccess -eq $true

            If($NotTitle -and $UpdateAccess -eq $true)
            {
                if ($Debuger) { Write-PmLog -Level Debug -Message "Set post search criteria: NotTitle = '$NotTitle'" }
                If($Update.Title -match $NotTitle)
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "UpdateAccess: false" }
                    $UpdateAccess = $false
                } #End If $Update.Title -notmatch $NotTitle
            } #End If $NotTitle -and $UpdateAccess -eq $true
            
            If($IgnoreUserInput -and $UpdateAccess -eq $true)
            {
                if ($Debuger) { Write-PmLog -Level Debug "Set post search criteria: CanRequestUserInput" }
                If($Update.InstallationBehavior.CanRequestUserInput -eq $true)
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "UpdateAccess: false" }
                    $UpdateAccess = $false
                } #End If $Update.InstallationBehavior.CanRequestUserInput -eq $true
            } #End If $IgnoreUserInput -and $UpdateAccess -eq $true

            If($IgnoreRebootRequired -and $UpdateAccess -eq $true) 
            {
                if ($Debuger) { Write-PmLog -Level Debug -Message "Set post search criteria: RebootBehavior" }
                If($Update.InstallationBehavior.RebootBehavior -ne 0)
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "UpdateAccess: false" }
                    $UpdateAccess = $false
                } #End If $Update.InstallationBehavior.RebootBehavior -ne 0    
            } #End If $IgnoreRebootRequired -and $UpdateAccess -eq $true

            If($UpdateAccess -eq $true)
            {
                if ($Debuger) { Write-PmLog -Level Debug -Message "Convert size" }
                Switch($Update.MaxDownloadSize)
                {
                    {[System.Math]::Round($_/1KB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1KB,0))+" KB"; break }
                    {[System.Math]::Round($_/1MB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1MB,0))+" MB"; break }  
                    {[System.Math]::Round($_/1GB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1GB,0))+" GB"; break }    
                    {[System.Math]::Round($_/1TB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1TB,0))+" TB"; break }
                    default { $size = $_+"B" }
                } #End Switch
            
                if ($Debuger) { Write-PmLog -Level Debug -Message "Convert KBArticleIDs" }
                If($Update.KBArticleIDs -ne "")    
                {
                    $KB = "KB"+$Update.KBArticleIDs
                } #End If $Update.KBArticleIDs -ne ""
                Else 
                {
                    $KB = ""
                } #End Else $Update.KBArticleIDs -ne ""
                
                If($ListOnly)
                {
                    $Status = ""
                    If($Update.IsDownloaded)    {$Status += "D"} else {$status += "-"}
                    If($Update.IsInstalled)     {$Status += "I"} else {$status += "-"}
                    If($Update.IsMandatory)     {$Status += "M"} else {$status += "-"}
                    If($Update.IsHidden)        {$Status += "H"} else {$status += "-"}
                    If($Update.IsUninstallable) {$Status += "U"} else {$status += "-"}
                    If($Update.IsBeta)          {$Status += "B"} else {$status += "-"} 
    
                    Add-Member -InputObject $Update -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME
                    Add-Member -InputObject $Update -MemberType NoteProperty -Name KB -Value $KB
                    Add-Member -InputObject $Update -MemberType NoteProperty -Name Size -Value $size
                    Add-Member -InputObject $Update -MemberType NoteProperty -Name Status -Value $Status
                    Add-Member -InputObject $Update -MemberType NoteProperty -Name X -Value 1
                    <#
                    $Update.PSTypeNames.Clear()
                    $Update.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
                    #>
                    $UpdateCollection += $Update
                } #End If $ListOnly
                Else
                {
                    $objCollectionUpdate.Add($Update) | Out-Null
                    $UpdatesExtraDataCollection.Add($Update.Identity.UpdateID,@{KB = $KB; Size = $size})
                } #End Else $ListOnly
            } #End If $UpdateAccess -eq $true
            
            $NumberOfUpdate++
        } #End Foreach $Update in $objResults.Updates
        
        $sMessage = '[1/' + $NumberOfStage + '] Get available updates completed.'
        if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
        {
            Write-PmLog -Level Information -Message $sMessage
        }
        else
        {
            Write-PmLog -Level Information -Message $sMessage
        }

        If($ListOnly)
        {
            $FoundUpdatesToDownload = $UpdateCollection.count
        } #End If $ListOnly
        Else
        {
            $FoundUpdatesToDownload = $objCollectionUpdate.count                
        } #End Else $ListOnly
        
        

        if ($Debuger) { Write-PmLog -Level Debug -Message "Found [$FoundUpdatesToDownload] Updates in post search criteria" }
        
        If($FoundUpdatesToDownload -eq 0)
        {
            $sMessage = 'There are no updates available.'
            if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
            {
                Write-PmLog -Level Information -Message $sMessage
            }
            else
            {
                Write-PmLog -Level Information -Message $sMessage
            }

            Return
        } #End If $FoundUpdatesToDownload -eq 0
        
        If($ListOnly)
        {
            if ($Debuger) { Write-PmLog -Level Debug -Message "Return only list of updates" }
            Return $UpdateCollection                
        } #End If $ListOnly

        <#
            End STAGE 1: Get updates list 
        #>        

        If(!$ListOnly) 
        {
            <#
                Start STAGE 2: Choose updates
            #>  
            
            if ($Debuger) { Write-PmLog -Level Debug -Message "STAGE 2: Choose updates"   }         
            $NumberOfUpdate = 1
            $logCollection = @()
            
            $objCollectionChoose = New-Object -ComObject "Microsoft.Update.UpdateColl"

            Foreach($Update in $objCollectionUpdate)
            {    
                $size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
                
                $sMessage = '[2/' + $NumberOfStage + '] Choose updates [' + $NumberOfUpdate + '/' + $FoundUpdatesToDownload + '] ' + $Update.Title + ' ' + $size
                if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                {
                    Write-PmLog -Level Information -Message $sMessage
                }
                else
                {
                    Write-PmLog -Level Information -Message $sMessage
                }
                
                if ($Debuger) { Write-PmLog -Level Debug -Message "Show update to accept: $($Update.Title)" }
                
                If($AcceptAll)
                {
                    $Status = "Accepted"

                    If($Update.EulaAccepted -eq 0)
                    { 
                        if ($Debuger) { Write-PmLog -Level Debug -Message "Accept Eula" }
                        $Update.AcceptEula() 
                    } #End If $Update.EulaAccepted -eq 0
            
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Add update to collection" }
                    $objCollectionChoose.Add($Update) | Out-Null
                } #End If $AcceptAll
                ElseIf($AutoSelectOnly)  
                {  
                    If($Update.AutoSelectOnWebsites)  
                    {  
                        $Status = "Accepted"  
                        If($Update.EulaAccepted -eq 0)  
                        {  
                            if ($Debuger) { Write-PmLog -Level Debug -Message "Accept Eula"   }
                            $Update.AcceptEula()  
                        } #End If $Update.EulaAccepted -eq 0 
  
                        if ($Debuger) { Write-PmLog -Level Debug -Message "Add update to collection"  }
                        $objCollectionChoose.Add($Update) | Out-Null  
                    } #End If $Update.AutoSelectOnWebsites 
                    Else  
                    {  
                        $Status = "Rejected"  
                    } #End Else $Update.AutoSelectOnWebsites
                } #End ElseIf $AutoSelectOnly
                Else
                {
                    If($pscmdlet.ShouldProcess($Env:COMPUTERNAME,"$($Update.Title)[$size]?")) 
                    {
                        $Status = "Accepted"
                        
                        If($Update.EulaAccepted -eq 0)
                        { 
                            if ($Debuger) { Write-PmLog -Level Debug -Message "Accept Eula" }
                            $Update.AcceptEula() 
                        } #End If $Update.EulaAccepted -eq 0
                
                        if ($Debuger) { Write-PmLog -Level Debug -Message "Add update to collection" }
                        $objCollectionChoose.Add($Update) | Out-Null 
                    } #End If $pscmdlet.ShouldProcess($Env:COMPUTERNAME,"$($Update.Title)[$size]?")
                    Else
                    {
                        $Status = "Rejected"
                    } #End Else $pscmdlet.ShouldProcess($Env:COMPUTERNAME,"$($Update.Title)[$size]?")
                } #End Else $AutoSelectOnly
                
                if ($Debuger) { Write-PmLog -Level Debug -Message "Add to log collection" }
                $log = New-Object PSObject -Property @{
                    Title = $Update.Title
                    KB = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].KB
                    Size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
                    Status = $Status
                    X = 2
                } #End PSObject Property
<#                
                $log.PSTypeNames.Clear()
                $log.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
#>                
                $logCollection += $log
                
                $NumberOfUpdate++
            } #End Foreach $Update in $objCollectionUpdate

            $sMessage = '[2/' + $NumberOfStage + '] Choose updates Completed.'
            if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
            {
                Write-PmLog -Level Information -Message $sMessage
            }
            else
            {
                Write-PmLog -Level Information -Message $sMessage
            }
            
            if ($Debuger) { Write-PmLog -Level Debug -Message "Show log collection" }
         
            $AcceptUpdatesToDownload = $objCollectionChoose.count
            if ($Debuger) { Write-PmLog -Level Debug -Message "Accept [$AcceptUpdatesToDownload] Updates to Download" }
            
            If($AcceptUpdatesToDownload -eq 0)
            {
                Return
            } #End If $AcceptUpdatesToDownload -eq 0    
                
            <#
                End STAGE 2: Choose updates 
            #>
            
            <#
                Start STAGE 3: Download updates 
            #>
            
            if ($Debuger) { Write-PmLog -Level Debug -Message "STAGE 3: Download updates" }
            $NumberOfUpdate = 1
            $objCollectionDownload = New-Object -ComObject "Microsoft.Update.UpdateColl" 

            Foreach($Update in $objCollectionChoose)
            {
                $sMessage = '[3/' + $NumberOfStage + '] Downloading updates [' + $NumberOfUpdate + '/' + $AcceptUpdatesToDownload + '] ' + $Update.Title + ' ' + $size
                Write-PmLog -Level Information -Message $sMessage

                if ($Debuger) { Write-PmLog -Level Debug -Message "Show update to download: $($Update.Title)" }
                if ($Debuger) { Write-PmLog -Level Debug -Message "Send update to download collection" }

                $objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
                $objCollectionTmp.Add($Update) | Out-Null
                    
                $Downloader         = $objSession.CreateUpdateDownloader() 
                $Downloader.Updates = $objCollectionTmp
                Try
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Try download update" }
                    $DownloadResult = $Downloader.Download()
                } #End Try
                Catch
                {
                    If($_ -match "HRESULT: 0x80240044")
                    {
                        $sMessage = "Your security policy don't allow a non-administator identity to perform this task"
                        if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                        {
                            Write-PmLog -Level Information -Message $sMessage
                        }
                        else
                        {
                            Write-PmLog -Level Information -Message $sMessage
                        }
                        
                    } #End If $_ -match "HRESULT: 0x80240044"
                    
                    Return
                } #End Catch 
                
                if ($Debuger) { Write-PmLog -Level Debug -Message "Check ResultCode" }
                Switch -exact ($DownloadResult.ResultCode)
                {
                    0   { $Status = "NotStarted" }
                    1   { $Status = "InProgress" }
                    2   { $Status = "Downloaded" }
                    3   { $Status = "DownloadedWithErrors" }
                    4   { $Status = "Failed" }
                    5   { $Status = "Aborted" }
                } #End Switch
                
                if ($Debuger) { Write-PmLog -Level Debug -Message "Add to log collection" }
                $log = New-Object PSObject -Property @{
                    Title = $Update.Title
                    KB = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].KB
                    Size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
                    Status = $Status
                    X = 3
                } #End PSObject Property
 <#               
                $log.PSTypeNames.Clear()
                $log.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
 #>               
                $logCollection += $log

                If($DownloadResult.ResultCode -eq 2)
                {
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Downloaded then send update to next stage" }
                    $objCollectionDownload.Add($Update) | Out-Null
                } #End If $DownloadResult.ResultCode -eq 2
                
                $NumberOfUpdate++
                
            } #End Foreach $Update in $objCollectionChoose

            $sMessage = '[3/' + $NumberOfStage + '] Downloading updates Completed.'
            if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
            {
                Write-PmLog -Level Information -Message $sMessage
            }
            else
            {
                Write-PmLog -Level Information -Message $sMessage
            }

            $ReadyUpdatesToInstall = $objCollectionDownload.count
            if ($Debuger) { Write-PmLog -Level Debug -Message "Downloaded [$ReadyUpdatesToInstall] Updates to Install" }
        
            If($ReadyUpdatesToInstall -eq 0)
            {
                Return
            } #End If $ReadyUpdatesToInstall -eq 0
        

            <#
                End STAGE 3: Download updates 
            #>
            
            If(!$DownloadOnly)
            {
                <#
                    Start STAGE 4: Install updates
                #>
                
                if ($Debuger) { Write-PmLog -Level Debug -Message "STAGE 4: Install updates" }
                $NeedsReboot = $false
                $NumberOfUpdate = 1
                
                #install updates    
                Foreach($Update in $objCollectionDownload)
                {   
                    
                    $sMessage = '[4/' + $NumberOfStage + '] Installing updates [' + $NumberOfUpdate + '/' + $ReadyUpdatesToInstall + '] ' + $Update.Title
                    if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                    {
                        Write-PmLog -Level Information -Message $sMessage
                    }
                    else
                    {
                        Write-PmLog -Level Information -Message $sMessage
                    }
                                       
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Show update to install: $($Update.Title)" }
                    
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Send update to install collection" }
                    $objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
                    $objCollectionTmp.Add($Update) | Out-Null
                    
                    $objInstaller = $objSession.CreateUpdateInstaller()
                    $objInstaller.Updates = $objCollectionTmp
                        
                    Try
                    {
                        if ($Debuger) { Write-PmLog -Level Debug -Message "Try install update" }
                        $InstallResult = $objInstaller.Install()
                    } #End Try
                    Catch
                    {
                        If($_ -match "HRESULT: 0x80240044")
                        {
                            
                            $sMessage = "Your security policy don't allow a non-administator identity to perform this task"
                            if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                            {
                                Write-PmLog -Level Information -Message $sMessage
                            }
                            else
                            {
                                Write-PmLog -Level Information -Message $sMessage
                            }
                        } #End If $_ -match "HRESULT: 0x80240044"
                        
                        Return
                    } #End Catch
                    
                    If(!$NeedsReboot) 
                    { 
                        if ($Debuger) { Write-PmLog -Level Debug -Message "Set instalation status RebootRequired" }
                        $NeedsReboot = $installResult.RebootRequired 
                    } #End If !$NeedsReboot
                    
                    Switch -exact ($InstallResult.ResultCode)
                    {
                        0   { $Status = "NotStarted"}
                        1   { $Status = "InProgress"}
                        2   { $Status = "Installed"}
                        3   { $Status = "InstalledWithErrors"}
                        4   { $Status = "Failed"}
                        5   { $Status = "Aborted"}
                    } #End Switch
                   
                    if ($Debuger) { Write-PmLog -Level Debug -Message "Add to log collection" }
                    $log = New-Object PSObject -Property @{
                        Title = $Update.Title
                        KB = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].KB
                        Size = $UpdatesExtraDataCollection[$Update.Identity.UpdateID].Size
                        Status = $Status
                        X = 4
                    } #End PSObject Property
<#                    
                    $log.PSTypeNames.Clear()
                    $log.PSTypeNames.Add('PSWindowsUpdate.WUInstall')
#>
                    $logCollection += $log

                    $sMessage = '[4/' + $NumberOfStage + '] Installing updates [' + $NumberOfUpdate + '/' + $ReadyUpdatesToInstall + '] ' + $Update.Title + ' - state: ' + $Status
                    if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                    {
                        Write-PmLog -Level Information -Message $sMessage
                    }
                    else
                    {
                        Write-PmLog -Level Information -Message $sMessage
                    }

                    $NumberOfUpdate++
                } #End Foreach $Update in $objCollectionDownload

                $sMessage = '[4/' + $NumberOfStage + '] Installing updates Completed.'
                if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                {
                    Write-PmLog -Level Information -Message $sMessage
                }
                else
                {
                    Write-PmLog -Level Information -Message $sMessage
                }
                
                If($NeedsReboot)
                {
                    If($AutoReboot)
                    {
                        Restart-ProComputer -TimeOut 300
                    } #End If $AutoReboot
                    ElseIf($IgnoreReboot)
                    {
                        $sMessage = "Reboot is required, but do it manually."
                        if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                        {
                            Write-PmLog -Level Information -Message $sMessage
                        }
                        else
                        {
                            Write-PmLog -Level Information -Message $sMessage
                        }
                    } #End Else $AutoReboot If $IgnoreReboot
                    Else
                    {
                        $Reboot = Read-Host "Reboot is required. Do it now ? [Y/N]"
                        if ($Debuger) { Write-PmLog -Level Debug -Message "Reboot is required. Do it now ? [Y/N] $Reboot" }
                        If($Reboot -eq "Y")
                        {
                            $sMessage = "Rebooting..."
                            if ($currentUserName -eq 'NT AUTHORITY\SYSTEM')
                            {
                                Write-PmLog -Level Information -Message $sMessage
                            }
                            else
                            {
                                Write-PmLog -Level Information -Message $sMessage
                            }
                            Restart-Computer -Force
                        } #End If $Reboot -eq "Y"
                        
                    } #End Else $IgnoreReboot    
                    
                } #End If $NeedsReboot

                <#
                    End STAGE 4: Install updates
                #>
            } #End If !$DownloadOnly
        } #End !$ListOnly
    } #End Process
    
    End
    {
        return $logCollection
    }        
}
