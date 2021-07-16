# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format.
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' property is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

Add-Type -AssemblyName System.Web

function Write-Log
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [Alias('LogPath')]
        [string]$Path='C:\Logs\PowerShellLog.log',
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Error","Warn","Info")]
        [string]$Level="Info",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoClobber
    )

    Begin
    {
        # Set VerbosePreference to Continue so that verbose messages are displayed.
        $VerbosePreference = 'Continue'
    }
    Process
    {
        
        # If the file already exists and NoClobber was specified, do not write to the log.
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
            }

        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            $NewLogFile = New-Item $Path -Force -ItemType File
            }

        else {
            # Nothing to see here yet.
            }

        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write message to error, warning, or verbose pipeline and specify $LevelText
        switch ($Level) {
            'Error' {
                Write-Error $Message
                $LevelText = 'ERROR:'
                }
            'Warn' {
                Write-Warning $Message
                $LevelText = 'WARNING:'
                }
            'Info' {
                Write-Verbose $Message
                $LevelText = 'INFO:'
                }
            }
        
        # Write log entry to $Path
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append
        ## also dump to console
        #$savedColor = $host.UI.RawUI.ForegroundColor 
        #$host.UI.RawUI.ForegroundColor = "DarkGreen"
        Write-Output  $message 
        #$host.UI.RawUI.ForegroundColor = $savedColor
    }
    End
    {
    }
}

function Process-JIT
{
    [CmdletBinding()]
    Param
    (
    )

    begin
    {
        #set the tenant variable in the configuration like M365x088487
        $tenant = $env:TENANTNAME;
        ##### Note: This Azure function is looking at the devlopment master requests...
        $jitRequestSite = "${env:JITREQUESTSITE}";
        $jitRequestWebUrl = "https://{0}.sharepoint.com{1}/" -f $tenant, $jitRequestSite;
        $JITSiteAdminRequestList = "Site Admin Requests” 
        # Total only five approved requests are serverd in an hour. 
        $MaxServePendingRequest = 4 
        $pendingRequestStatu2Check = "Pending"
        $activeRequestStatu2Check  = "Active"
        $approvedRequestStatu2Check = "Approved"
        $siteUrl2Prefix = "https://{0}.sharepoint.com/sites/" -f $tenant;
    }
    process
    {
        try
        {
            # the following check is only to debu on Power Shell ISE
            if ( $env:certificateBase64Encode -ne $null)
            {
                $certificateBase64Encode = $env:certificateBase64Encode
            }
            else
            {
                # get the PFX secret from the key vault
                $kvSecret = Get-AzKeyVaultSecret -VaultName $env:KeyVaultName -Name $env:KeyVaultSecretName
                $certificateBase64Encode = '';
                $ssPtr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($kvSecret.SecretValue)
                try {
                    $certificateBase64Encode = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ssPtr)
                } finally {
                    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ssPtr)
                }
            }
            Write-Log -Path $LogFileName $("Connecting to {0}" -f $jitRequestWebUrl);
            Write-Log -Path $LogFileName $("webUrl  {0}" -f $jitRequestWebUrl );
            # Using Splat to convert
            $HashArguments = @{
                    Url                      = $jitRequestWebUrl
                    ClientId                 = $env:GRAPH_APP_ID
                    CertificateBase64Encoded = $certificateBase64Encode
                    Tenant                   = $("{0}.onmicrosoft.com" -f  $tenant)
            }
            $jitRequestSitesConnection = Connect-PnPOnline @HashArguments  -ReturnConnection
            #
            # Get Context to use later
            #
            $jitRequestSitesContext = Get-PnPContext

            Write-Log -Path $LogFileName $("Connected to {0}" -f $jitRequestWebUrl);
            Write-Log -Path $LogFileName $("Getting items from {0} for reqStatus={1} OR reqStatus={2}." -f $JITSiteAdminRequestList,  $pendingRequestStatu2Check, $activeRequestStatu2Check);
            $HashArguments = @{
                    Connection      = $jitRequestSitesConnection
                    List            = $JITSiteAdminRequestList 
            }
            # only get ID, reqStatus and reqSysStatus
            [array]$items = Get-PnPListItem @HashArguments | `
                    %{New-Object psobject -Property `
                    @{ 
                        Id                          = $_.Id; 
                        reqStatus                   = $_["reqStatus"];
                        reqActivateTime             = $_["reqActivateTime"];
                        reqExpireTimeMin            = $_["reqExpireTimeMin"];
	                 }} | `
                     select ID, 
                            reqStatus,
                            reqActivateTime, 
                            reqExpireTimeMin    
            Disconnect-PnPOnline -Connection $jitRequestSitesConnection
            
            $today = Get-Date
            $today = $today.ToUniversalTime();
            #$minutesPassed = ($today - $item.reqActiveStartTime).Minutes


            [array]$PendingFilteredItems   = $items | Where-Object {$_.reqStatus -contains $pendingRequestStatu2Check}  
            [array]$ApprovedFilteredItems  = $items | Where-Object {$_.reqStatus -contains $approvedRequestStatu2Check} 
            [array]$ActiveFilteredItems    = $items | Where-Object {(($_.reqStatus -contains $activeRequestStatu2Check) -and ( ($today - $_.reqActivateTime).TotalMinutes -gt $_.reqExpireTimeMin))}   


            Write-Log -Path $LogFileName $("There are {0} {1}, {2} {3} and {4} {5} RequestStatus. Total {6} items for {7} list on {8}." -f $PendingFilteredItems.Count, 
                                $pendingRequestStatu2Check, 
                                $ApprovedFilteredItems.Count, 
                                $approvedRequestStatu2Check,$ActiveFilteredItems.Count , 
                                $activeRequestStatu2Check, $items.Count, 
                                $JITSiteAdminRequestList,
                                $jitRequestWebUrl
                                );

            [array]$CombinedFilteredItems = $PendingFilteredItems + $ActiveFilteredItems + $ApprovedFilteredItems| Sort-Object -Property Id

            if ( ($CombinedFilteredItems -ne $null) -and ($CombinedFilteredItems.Count -gt 0) )
            {
                foreach ($aCombinedFilteredItem in $CombinedFilteredItems)
                {
                    $HashArguments = @{
                            Url                      = $jitRequestWebUrl
                            ClientId                 = $env:GRAPH_APP_ID
                            CertificateBase64Encoded = $certificateBase64Encode
                            Tenant                   = $("{0}.onmicrosoft.com" -f  $tenant)
                    }

                    $jitRequestSitesConnection = Connect-PnPOnline @HashArguments  -ReturnConnection
                    $id2Fetch = $aCombinedFilteredItem.Id
                    # technically there will be only one item in the collection. Get everything for this item.
                    $item = Get-PnPListItem -List $JITSiteAdminRequestList -Id $id2Fetch | `
                        %{New-Object psobject -Property `
                        @{ 
                            Id                          = $_.Id; 
                            Title                       = $_["Title"];
                            reqJustification            = $_["reqJustification"]; 
                            reqStatus                   = $_["reqStatus"];
                            reqSysStatus                = $_["reqSysStatus"];
                            reqActivateTime             = $_["reqActivateTime"];
                            reqExpiryTime               = $_["reqExpiryTime"];
                            reqExpireTimeMin            = $_["reqExpireTimeMin"];
                            reqApprovers                = $_["reqApprovers"];
                            Author                      = $_["Author"];
                            Created                     = $_["Created"];
                            Editor                      = $_["Editor"];
                            Modified                    = $_["Modified"];
	                            }} | `
                            select ID, 
                                Title,            
                                reqJustification,
                                reqStatus,        
                                reqSysStatus,
                                reqActivateTime,
                                reqExpiryTime,
                                reqExpireTimeMin,
                                reqApprovers,
                                Author,           
                                Created,          
                                Editor,           
                                Modified
                    Disconnect-PnPOnline -Connection $jitRequestSitesConnection
                    Write-Log -Path $LogFileName $("Got item properties for Id = {0} Title = {1}." -f $aCombinedFilteredItem.Id, $item.Title);
                    # manupulate the absolute url from the Title field as relative url i.e. https://[tenant-name]..sharepoint.com/sites/[whatever-Title-value]
                    $requestedSiteUrl = $siteUrl2Prefix + $item.Title;
                    if ( $requestedSiteUrl[$requestedSiteUrl.Length-1] -eq '/' )
                    {
                        # found the last character as / remove the last character
                        $requestedSiteUrl = $requestedSiteUrl.Substring(0,$requestedSiteUrl.Length - 1)
                    }

                    $user2AddRemove = $item.Author.LookupValue

                    # check if the status is pending and system status is NOT TRIGGER-ADMIN-UPDATED
                    if ( ($item -ne $null) -and ($item.reqStatus.ToUpper().Contains("PENDING") -eq $true) -and ($item.reqSysStatus.ToUpper().Contains("TRIGGER-ADMIN-UPDATED") -eq $false) )
                    {
                        # Connect to the requested site. get SCAs of it.
                        $HashArguments = @{
                                Url                      = $requestedSiteUrl
                                ClientId                 = $env:GRAPH_APP_ID
                                CertificateBase64Encoded = $certificateBase64Encode
                                Tenant                   = $("{0}.onmicrosoft.com" -f  $tenant)
                        }
                        Write-Log -Path $LogFileName $("Connecting to {0}" -f $requestedSiteUrl);
                        $requestSitesConnection = Connect-PnPOnline @HashArguments  -ReturnConnection
                        Write-Log -Path $LogFileName $("Connected to {0}" -f $requestedSiteUrl);
                        
                        # get all SCAs 
                        Write-Log -Path $LogFileName $("Getting SCAs for {0}" -f  $item.reqURL);
                        $scaColl = Get-PnPSiteCollectionAdmin -Connection $requestSitesConnection -ErrorAction Stop
                        Write-Log -Path $LogFileName $("Got {0} SCAs for {1}" -f  $scaColl.Count, $requestedSiteUrl);


                        $listOfSCAforURL = New-Object System.Collections.ArrayList
                        $scaIdx = 1;
                        foreach($sca in $scaColl)  
                        {          
                            Write-Log -Path $LogFileName $(" {0} = {1}" -f $scaIdx++, $sca.LoginName);
                            # sample values in the $sca variable.
                            # Id Title                         LoginName                                                                     Email                                        
                            # -- -----                         ---------                                                                     -----                                        
                            # 16 Surti, Pankaj   .             i:0#.f|membership|pankaj.surti@contoso.com                                                                                     
                            #  6 JITSiteAdminAccess-dev Owners c:0o.c|federateddirectoryclaimprovider|f4c66d17-e431-4430-8ff5-180562cd9ce9_o JITSiteAdminAccess-dev@CONTOSO.onmicrosoft.com
                            if ( $sca.LoginName.ToLower().Contains("federateddirectoryclaimprovider") -eq $true )
                            {

                                # found the Office 365 group
                                $ownersO365Grp = Get-PnPMicrosoft365GroupOwners -Identity $sca.Email.Split('@')[0]
                                # now iterate over the members to fill the make SCA list
                                foreach($anOwner in $ownersO365Grp)  
                                {
                                    # sample values in the $anOwner variable, only take Email values.
                                    #Id                                   DisplayName                                  UserPrincipalName                       Email                 
                                    #--                                   -----------                                  -----------------                       -----                 
                                    #7e570ddc-fabb-4aa3-9832-aa04b162874e Surti, Pankaj                                Pankaj.Surti@contoso.com                Pankaj.Surti@contoso.com
                                    #69693091-4083-4709-9964-efbee076aecd Douglas, Sean T.                             Sean.Douglas@contoso.com                Sean.Douglas@contoso.com   
                                    if ($anOwner.Email.Length -gt 0 )
                                    {
                                        $listOfSCAforURL.Add($anOwner.Email)
                                    }
                                }
                            }
                            elseif ( $sca.LoginName.ToLower().Contains("membership") -eq $true )
                            {
                                if ( $sca.Email.Length -gt 0 )
                                {
                                    $listOfSCAforURL.Add($sca.Email)
                                }
                            }
                        }
                        # disconnect now
                        Disconnect-PnPOnline -Connection $requestSitesConnection

                        # now add the SCA list to the list item of the list $JITSiteAdminRequestList.
                        if ($listOfSCAforURL.Count -gt 0)
                        {

                            $HashArguments = @{
                                    Url                      = $jitRequestWebUrl
                                    ClientId                 = $env:GRAPH_APP_ID
                                    CertificateBase64Encoded = $certificateBase64Encode
                                    Tenant                   = $("{0}.onmicrosoft.com" -f  $tenant)
                            }
                            $jitRequestSitesConnection = Connect-PnPOnline @HashArguments  -ReturnConnection
                            Write-Log -Path $LogFileName $("Setting reqApprovers to {0} SCAs AND reqSysStatus to 'TRIGGER-ADMIN-UPDATED'." -f $listOfSCAforURL.Count);
                            $arrOfSCAs = $listOfSCAforURL.ToArray()
                            Write-Log -Path $LogFileName $("Set reqSysStatus=TRIGGER-ADMIN-UPDATED");
                            $fakeNotUsedItem = Set-PnPListItem -Connection $jitRequestSitesConnection -List $JITSiteAdminRequestList -Identity $item.Id -Values @{ 
                                                        "reqApprovers"= $arrOfSCAs; 
                                                        "reqSysStatus" = "TRIGGER-ADMIN-UPDATED"; 
                                               }
                            Write-Log -Path $LogFileName $("Successfully set of reqApprovers to {0} SCAs AND reqSysStatus to 'TRIGGER-ADMIN-UPDATED'." -f $listOfSCAforURL.Count);
                            Disconnect-PnPOnline -Connection $jitRequestSitesConnection



                        }
                    }
                    if ( ($item -ne $null) -and ($item.reqStatus.ToUpper().Contains("APPROVED") -eq $true) )
                    {
                        $user2AddOrremove
                        Write-Log -Path $LogFileName $("Setting owner {0} to {1}" -f $user2AddRemove, $requestedSiteUrl);
                        # Connect to the requested site.
                        $HashArguments = @{
                                Url                      = $requestedSiteUrl
                                ClientId                 = $env:GRAPH_APP_ID
                                CertificateBase64Encoded = $certificateBase64Encode
                                Tenant                   = $("{0}.onmicrosoft.com" -f  $tenant)
                        }
                        Write-Log -Path $LogFileName $("Connecting to {0}" -f $requestedSiteUrl);
                        $requestSitesConnection = Connect-PnPOnline @HashArguments  -ReturnConnection
                        Write-Log -Path $LogFileName $("Connected to {0}" -f $requestedSiteUrl);
                        $temp = Set-PnPTenantSite -Url $requestedSiteUrl -Owners $user2AddRemove -Connection $requestSitesConnection -ErrorAction Stop
                        Write-Log -Path $LogFileName $("Owner {0} is set to {1}" -f $user2AddRemove, $requestedSiteUrl);
                        Disconnect-PnPOnline -Connection $requestSitesConnection

                        # get the current time and date.

                        $HashArguments = @{
                                Url                      = $jitRequestWebUrl
                                ClientId                 = $env:GRAPH_APP_ID
                                CertificateBase64Encoded = $certificateBase64Encode
                                Tenant                   = $("{0}.onmicrosoft.com" -f  $tenant)
                        }
                        $jitRequestSitesConnection = Connect-PnPOnline @HashArguments  -ReturnConnection
                        $currentDate = Get-Date
                        $timeSpan    = New-TimeSpan -Minutes $item.reqExpireTimeMin
                        $expireTime  = $currentDate + $timeSpan
                        Write-Log -Path $LogFileName $("Set reqSysStatus=TRIGGER-USER-ACTIVE, reqActiveStartTime={0} and reqStatus=Active" -f $currentDate);
                        $fakeNotUsedItem = Set-PnPListItem -Connection $jitRequestSitesConnection -List $JITSiteAdminRequestList -Identity $item.Id -Values @{ 
                                            "reqSysStatus"          = "TRIGGER-USER-ACTIVE"; 
                                            "reqStatus"             = "Active"; 
                                            "reqAbsoluteSiteUrl"    = $requestedSiteUrl;
                                            "reqActivateTime"       = $currentDate; 
                                            "reqExpiryTime"         = $expireTime;
                                            }
                        Disconnect-PnPOnline -Connection $jitRequestSitesConnection

                    }
                    elseif ( ($item -ne $null) -and ($item.reqStatus.ToUpper().Contains("ACTIVE") -eq $true) )
                    {
                        Write-Log -Path $LogFileName $("Today is {0} Expire {1} minutes passed condition macthed for actual {2} minutes." -f $today, $item.reqExpireTimeLimitInMinutes, ($today - $item.reqActivateTime).TotalMinutes);
                        Write-Log -Path $LogFileName $("Removing owner {0} to {1}" -f $user2AddRemove, $requestedSiteUrl);

                        #$usrWithClaims = "i:0#.f|membership|" + $item.Author.LookupValue
                        # Connect to the requested site.
                        $HashArguments = @{
                                Url                      = $requestedSiteUrl
                                ClientId                 = $env:GRAPH_APP_ID
                                CertificateBase64Encoded = $certificateBase64Encode
                                Tenant                   = $("{0}.onmicrosoft.com" -f  $tenant)
                        }
                        Write-Log -Path $LogFileName $("Connecting to {0}" -f $requestedSiteUrl);
                        $requestSitesConnection = Connect-PnPOnline @HashArguments  -ReturnConnection
                        Write-Log -Path $LogFileName $("Connected to {0}" -f $requestedSiteUrl);
                        Remove-PnPSiteCollectionAdmin -Owners $user2AddRemove -Connection $requestSitesConnection -ErrorAction Stop
                        Write-Log -Path $LogFileName $("Owner {0} is REMOVED from {1}" -f $user2AddRemove, $requestedSiteUrl);
                        Disconnect-PnPOnline -Connection $requestSitesConnection



                        $HashArguments = @{
                                Url                      = $jitRequestWebUrl
                                ClientId                 = $env:GRAPH_APP_ID
                                CertificateBase64Encoded = $certificateBase64Encode
                                Tenant                   = $("{0}.onmicrosoft.com" -f  $tenant)
                        }
                        Write-Log -Path $LogFileName $("Set reqSysStatus=TRIGGER-USER-REMOVED and reqStatus=Completed");
                        $jitRequestSitesConnection = Connect-PnPOnline @HashArguments  -ReturnConnection
                        $fakeNotUsedItem = Set-PnPListItem -Connection $jitRequestSitesConnection -List $JITSiteAdminRequestList -Identity $item.Id -Values @{ 
                                            "reqSysStatus"          = "TRIGGER-USER-REMOVED"; 
                                            "reqStatus"             = "Completed"; }
                        Disconnect-PnPOnline -Connection $jitRequestSitesConnection

                    }
                    else
                    {
                        # this should never happen log and get out the loop
                        Write-Log -Path $LogFileName $("Id = {0} NOTE FOUND**** or status {1} not found." -f $id2Fetch, $item.reqStatus );
                    }

                }
            }
            else
            {
                # there is no need of the log file, simply delete it.
                Remove-Item -Path $LogFileName -Force

            }

        }
        catch
        {
            # https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-exceptions?view=powershell-7.1
            $ErrorMessage = $_.Exception | Out-String
            Write-Log -Path $LogFileName $("Exception {0}" -f $ErrorMessage );
            $InvocationInfo = $PSItem.InvocationInfo | Out-String
            Write-Log -Path $LogFileName $("InvocationInfo {0}" -f $InvocationInfo );
            $StackTrace = $PSItem.Exception.StackTrace | Out-String
            Write-Log -Path $LogFileName $("StackTrace {0}" -f $StackTrace );

            $HashArguments = @{
                    Url                      = $jitRequestWebUrl
                    ClientId                 = $env:GRAPH_APP_ID
                    CertificateBase64Encoded = $certificateBase64Encode
                    Tenant                   = $("{0}.onmicrosoft.com" -f  $tenant)
            }
            $jitRequestSitesConnection = Connect-PnPOnline @HashArguments  -ReturnConnection
            Write-Log -Path $LogFileName $("Set reqSysStatus=ERROR and reqStatus=Error");
            $fakeNotUsedItem = Set-PnPListItem -Connection $jitRequestSitesConnection -List $JITSiteAdminRequestList -Identity $item.Id -Values @{ 
                                "reqSysStatus"       = "ERROR"; 
                                "reqStatus"             = "Error";
                                "reqError" = $InvocationInfo + $StackTrace}
            Disconnect-PnPOnline -Connection $jitRequestSitesConnection
        }
    }
}

function MoveLogFilesToBlobContainer
{
    $jitRequestSitesConnectionectionString = $env:AzureWebJobsStorage;
    $containerName = $env:LOG_FILES_STORAGE_CONTAINER
    if ( $env:LOG_FILES_STORAGE_CONTAINER -ne $null )
    {
        $storageContainer = New-AzStorageContext -ConnectionString $jitRequestSitesConnectionectionString | Get-AzStorageContainer  -Name $containerName
        #Write-Output $storageContainer
        Get-ChildItem $env:LOG_FILE_PATH -Filter JIT-Log*.txt | 
        Foreach-Object {
            $blobNameWithFolder = $("{0}/{1}" -f (Get-Date -Format "yyyy-MM-dd"), $_.Name)
            Write-Output $("Move {0} to {1} Blob Container AS BlobName {2}." -f $_.FullName, $storageContainer.Name, $blobNameWithFolder)
            #Write-Output $("Moved {0} to {1} Blob Container." -f $_.FullName, $storageContainer.Name)
            Set-AzStorageBlobContent -File $_.FullName `
                -Container $storageContainer.Name `
                -Blob $blobNameWithFolder `
                -Context $storageContainer.Context -Force
            Remove-Item -Path $_.FullName -Force
        }
    }
}


# first process the new requests
$LogFileName = $("{0}\JIT-Log-{1}.txt" -f $env:LOG_FILE_PATH , (Get-Date -Format "yyyy-MM-dd-HH-mm-ss"))

Process-JIT 


Write-Output $("Call MoveLogFilesToBlobContainer" )
MoveLogFilesToBlobContainer
Write-Output $("After Call MoveLogFilesToBlobContainer" )


