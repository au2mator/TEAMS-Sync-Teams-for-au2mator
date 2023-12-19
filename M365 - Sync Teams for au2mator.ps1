
## Environment

[string]$LogPath = "C:\_SCOworkingDir\TFS\PS-Services\M365 - Sync Teams for au2mator"
[string]$LogfileName = "Sync Teams for au2mator"


$Modules = @("SQLserver") #$Modules = @("ActiveDirectory", "SharePointPnPPowerShellOnline")


$StagingDatabase = "au2matorHelp"
$StagingServer = "Demo01"


Function Write-au2matorLog {
    [CmdletBinding()]
    param
    (
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR')]
        [string]$Type,
        [string]$Text
    )
       
    # Set logging path
    if (!(Test-Path -Path $logPath)) {
        try {
            $null = New-Item -Path $logPath -ItemType Directory
            Write-Verbose ("Path: ""{0}"" was created." -f $logPath)
        }
        catch {
            Write-Verbose ("Path: ""{0}"" couldn't be created." -f $logPath)
        }
    }
    else {
        Write-Verbose ("Path: ""{0}"" already exists." -f $logPath)
    }
    [string]$logFile = '{0}\{1}_{2}.log' -f $logPath, $(Get-Date -Format 'yyyyMMdd'), $LogfileName
    $logEntry = '{0}: <{1}> <{2}> <{3}> {4}' -f $(Get-Date -Format dd.MM.yyyy-HH:mm:ss), $Type, $RequestId, $Service, $Text
    Add-Content -Path $logFile -Value $logEntry
}


Function Sync-Table_TeamsList ($Object) {
    $Object | add-member -NotePropertyName Status -NotePropertyValue "1"
    # Check if Table exists
    $TableName = "Teams-TeamsList"

    if ((Read-SqlTableData  -ServerInstance $StagingServer -DatabaseName $StagingDatabase -SchemaName 'dbo' -TableName $TableName -ErrorAction SilentlyContinue).count -eq 0) {
        $Object | select-Object -Property ID, Displayname, Visibility, MailNickName, Description, Status    | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName "Teams-TeamsList" -SchemaName "dbo" -Force
    }



    $QueryCheck = "Select ID from [$TableName] where ID = '$($Object.ID)'"
    $QueryUpdate = "USE [$StagingDatabase]
    GO

    UPDATE [dbo].[$TableName]
    SET [ID] = '$($Object.ID)'
        ,[DisplayName] = '$($Object.DisplayName)'
        ,[Visibility] = '$($Object.Visibility)'
        ,[MailNickName] = '$($Object.MailNickName)'
        ,[Description] = '$($Object.Description)'
        ,[Status] = '$($Object.Status)'
    WHERE ID = '$($Object.ID)'
    GO"

    


    if (Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryCheck) {
        # Update
        Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryUpdate

    }
    else {
        #Insert
        $Object | select-Object -Property ID, Displayname, Visibility, MailNickName, Description, Status     | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName "Teams-TeamsList" -SchemaName "dbo"
    }

}

Function Sync-Table_ChannelList ($Object, $TeamID) {
    $Object | add-member -NotePropertyName TeamID -NotePropertyValue $TeamID 
    $Object | add-member -NotePropertyName Status -NotePropertyValue "1"
    
    # Check if Table exists
    $TableName = "Teams-ChannelList"

    if ((Read-SqlTableData  -ServerInstance $StagingServer -DatabaseName $StagingDatabase -SchemaName 'dbo' -TableName $TableName -ErrorAction SilentlyContinue).count -eq 0) {
        
        $Object | select-Object -Property ID, Displayname, Description, TeamID, Status  | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName $TableName -SchemaName "dbo" -Force
    }


    $QueryCheck = "Select ID from [$TableName] where ID = '$($Object.ID)'"
    $QueryUpdate = "USE [$StagingDatabase]
    GO

    UPDATE [dbo].[$TableName]
    SET [ID] = '$($Object.id)'
        ,[DisplayName] = '$($Object.DisplayName)'
        ,[Description] = '$($Object.Description)'
        ,[TeamID] = '$($TeamID)'
        ,[Status] = '$($Object.Status)'
    WHERE ID = '$($Object.id)'
    GO"

    if (Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryCheck) {
        # Update
        Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryUpdate

    }
    else {
        #Insert

        $Object | select-Object -Property ID, Displayname, Description, TeamID, Status  | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName $TableName -SchemaName "dbo"
    }

}


Function Sync-Table_MemberList ($Object, $TeamID) {
    $Object | add-member -NotePropertyName TeamID -NotePropertyValue $TeamID 
    $Object | add-member -NotePropertyName Status -NotePropertyValue "1"
    if ($Member.roles -eq "Owner")
    {$Object | add-member -NotePropertyName Role -NotePropertyValue "Owner"}
    else {
        $Object | add-member -NotePropertyName Role -NotePropertyValue "Member"
    }
    # Check if Table exists
    $TableName = "Teams-MemberList"

    if ((Read-SqlTableData  -ServerInstance $StagingServer -DatabaseName $StagingDatabase -SchemaName 'dbo' -TableName $TableName -ErrorAction SilentlyContinue).count -eq 0) {
        
        $Object | select-Object -Property UserID, email, displayName, Role, TeamID, Status  | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName $TableName -SchemaName "dbo" -Force
    }

    
    $QueryCheck = "Select UserID from [$TableName] where UserID = '$($Object.UserID)'"
    $QueryUpdate = "USE [$StagingDatabase]
    GO

    UPDATE [dbo].[$TableName]
    SET [UserID] = '$($Object.UserID)'
        ,[email] = '$($Object.email)'
        ,[displayName] = '$($Object.displayName)'
        ,[Role] = '$($Object.Role)'
        ,[TeamID] = '$($TeamID)'
        ,[Status] = '$($Object.Status)'
    WHERE UserID = '$($Object.UserID)'
    GO"

    if (Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryCheck) {
        # Update
        Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryUpdate

    }
    else {
        #Insert
        
        $Object | select-Object -Property UserID, email, displayName, Role, TeamID, Status  | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName $TableName -SchemaName "dbo" -Force
    }

}

Function Reset-Deleted ($TableName) {
    $QueryDeleted = "USE [$StagingDatabase]
    GO

    UPDATE [dbo].[$TableName]
    SET [Status] = '0'
    GO"

    Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryDeleted



}


foreach ($Module in $Modules) {
    if (Get-Module -ListAvailable -Name $Module) {
        Write-au2matorLog -Type INFO -Text "Module is already installed:  $Module"        
    }
    else {
        Write-au2matorLog -Type INFO -Text "Module is not installed, try simple method:  $Module"
        try {

            Install-Module $Module -Force -Confirm:$false
            Write-au2matorLog -Type INFO -Text "Module was installed the simple way:  $Module"

        }
        catch {
            Write-au2matorLog -Type INFO -Text "Module is not installed, try the advanced way:  $Module"
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Install-PackageProvider -Name NuGet  -MinimumVersion 2.8.5.201 -Force
                Install-Module $Module -Force -Confirm:$false
                Write-au2matorLog -Type INFO -Text "Module was installed the advanced way:  $Module"

            }
            catch {
                Write-au2matorLog -Type ERROR -Text "could not install module:  $Module"
            }
        }
    }
    Write-au2matorLog -Type INFO -Text "Import Module:  $Module"
    Import-module $Module
}

[string]$CredentialStorePath = "C:\_SCOworkingDir\TFS\PS-Services\CredentialStore" #see for details: https://click.au2mator.com/PSCreds/?utm_source=github&utm_medium=social&utm_campaign=M365_SyncTeams&utm_content=PS1

$GraphAPICred_File = "TeamsCreds.xml"
$GraphAPICred = Import-CliXml -Path (Get-ChildItem -Path $CredentialStorePath -Filter $GraphAPICred_File).FullName
$clientId = $GraphAPICred.clientId
$clientSecret = $GraphAPICred.clientSecret
$tenantName = $GraphAPICred.tenantName



# Make Sure Database exists
Write-au2matorLog -Type INFO -Text "Make sure Database exists"
try {
    Get-SqlDatabase -Name $StagingDatabase -ServerInstance $StagingServer -ErrorAction Stop 
}
catch {
    $sql = "
    CREATE DATABASE $StagingDatabase
    "
    Invoke-SqlCmd -ServerInstance $StagingServer -Query $sql
    
}



# create variable with SQL to execute
Write-au2matorLog -Type INFO -Text "Connect to Microsoft Teams"





try {
    Write-au2matorLog -Type INFO -Text "Try to connect to Microsoft Teams"
    
    $tokenBody = @{  
        Grant_Type    = "client_credentials"  
        Scope         = "https://graph.microsoft.com/.default"  
        Client_Id     = $clientId  
        Client_Secret = $clientSecret  
    }   
  
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $tokenBody  

    $headers = @{
        "Authorization" = "Bearer $($tokenResponse.access_token)"
        "Content-type"  = "application/json"
    }


try {
    $URL = "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"  
    $AllTeams = (Invoke-RestMethod -Headers $headers -Uri $URL -Method GET).value 

    Reset-Deleted -TableName "Teams-TeamsList"
    Reset-Deleted -TableName "Teams-ChannelList"
    Reset-Deleted -TableName "Teams-MemberList"
    foreach ($Team in $AllTeams) {
        Write-au2matorLog -Type INFO -Text "Working with Teams: $($Team.DisplayName))"
        Sync-Table_TeamsList -Object $Team 
    
    
        # Take care about Channels
        $URL = "https://graph.microsoft.com/v1.0/teams/$($Team.id)/channels"  
    
        $AllChannels = (Invoke-RestMethod -Headers $headers -Uri $URL -Method GET).value 
        foreach ($Channel in $AllChannels) {
            Write-au2matorLog -Type INFO -Text "Working with Channel: $($Channel.DisplayName))"
            Sync-Table_ChannelList -Object $Channel -TeamID $Team.ID
        }

        $URL = "https://graph.microsoft.com/v1.0/teams/$($Team.id)/members"  
        $AllMembers = (Invoke-RestMethod -Headers $headers -Uri $URL -Method GET).value 
        foreach ($Member in $AllMembers) {
            Write-au2matorLog -Type INFO -Text "Working with User: $($Member.User))"
            Sync-Table_MemberList -Object $Member -TeamID $Team.ID
        }


    

}
}
catch {
    
}





}
catch {
    Write-au2matorLog -Type ERROR -Text "Error on connecting to Microsoft Teams, Error: $Error"

}




