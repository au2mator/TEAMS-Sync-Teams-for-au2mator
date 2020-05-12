
## Environment

[string]$LogPath = "C:\_SCOworkingDir\TFS\PS-Services\M365 - Sync Teams for au2mator"
[string]$LogfileName = "Sync Teams for au2mator"


$TeamsAdminUser = "TeamsAdmin@tenant.onmicrosoft.com"
$TeamsAdminPW = "YourPassword"

$StagingDatabase="au2matorHelp"
$StagingServer="Demo01"


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


Function Sync-Table_TeamsList ($Object)
{
    $Object | add-member -NotePropertyName Status -NotePropertyValue "1"
    # Check if Table exists
    $TableName="Teams-TeamsList"

    if ((Read-SqlTableData  -ServerInstance $StagingServer -DatabaseName $StagingDatabase -SchemaName 'dbo' -TableName $TableName -ErrorAction SilentlyContinue).count -eq 0)
    {
        $Object | select-Object -Property GroupID, Displayname, Visibility, Archived, MailNickName, Description, Status    | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName "Teams-TeamsList" -SchemaName "dbo" -Force
    }



    $QueryCheck="Select GroupID from [$TableName] where GroupID = '$($Object.groupid)'"
    $QueryUpdate="USE [$StagingDatabase]
    GO

    UPDATE [dbo].[$TableName]
    SET [GroupId] = '$($Object.groupid)'
        ,[DisplayName] = '$($Object.DisplayName)'
        ,[Visibility] = '$($Object.Visibility)'
        ,[Archived] = '$($Object.Archived)'
        ,[MailNickName] = '$($Object.MailNickName)'
        ,[Description] = '$($Object.Description)'
        ,[Status] = '$($Object.Status)'
    WHERE GroupID = '$($Object.groupid)'
    GO"

    


    if (Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryCheck)
    {
        # Update
        Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryUpdate

    }
    else {
        #Insert
        $Object | select-Object -Property GroupID, Displayname, Visibility, Archived, MailNickName, Description, Status     | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName "Teams-TeamsList" -SchemaName "dbo"
    }

}

Function Sync-Table_ChannelList ($Object, $TeamID)
{
    $Object | add-member -NotePropertyName TeamID -NotePropertyValue $TeamID 
    $Object | add-member -NotePropertyName Status -NotePropertyValue "1"
    
    # Check if Table exists
    $TableName="Teams-ChannelList"

    if ((Read-SqlTableData  -ServerInstance $StagingServer -DatabaseName $StagingDatabase -SchemaName 'dbo' -TableName $TableName -ErrorAction SilentlyContinue).count -eq 0)
    {
        
        $Object | select-Object -Property ID, Displayname, Description, TeamID, Status  | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName $TableName -SchemaName "dbo" -Force
    }


    $QueryCheck="Select ID from [$TableName] where ID = '$($Object.ID)'"
    $QueryUpdate="USE [$StagingDatabase]
    GO

    UPDATE [dbo].[$TableName]
    SET [ID] = '$($Object.id)'
        ,[DisplayName] = '$($Object.DisplayName)'
        ,[Description] = '$($Object.Description)'
        ,[TeamID] = '$($TeamID)'
        ,[Status] = '$($Object.Status)'
    WHERE ID = '$($Object.id)'
    GO"

    if (Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryCheck)
    {
        # Update
        Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryUpdate

    }
    else {
        #Insert

        $Object | select-Object -Property ID, Displayname, Description, TeamID, Status  | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName $TableName -SchemaName "dbo"
    }

}


Function Sync-Table_MemberList ($Object, $TeamID)
{
    $Object | add-member -NotePropertyName TeamID -NotePropertyValue $TeamID 
    $Object | add-member -NotePropertyName Status -NotePropertyValue "1"
    # Check if Table exists
    $TableName="Teams-MemberList"

    if ((Read-SqlTableData  -ServerInstance $StagingServer -DatabaseName $StagingDatabase -SchemaName 'dbo' -TableName $TableName -ErrorAction SilentlyContinue).count -eq 0)
    {
        
        $Object | select-Object -Property UserID, User, Name, Role, TeamID, Status  | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName $TableName -SchemaName "dbo" -Force
    }


    $QueryCheck="Select UserID from [$TableName] where UserID = '$($Object.UserID)'"
    $QueryUpdate="USE [$StagingDatabase]
    GO

    UPDATE [dbo].[$TableName]
    SET [UserID] = '$($Object.UserID)'
        ,[User] = '$($Object.User)'
        ,[Name] = '$($Object.Name)'
        ,[Role] = '$($Object.Role)'
        ,[TeamID] = '$($TeamID)'
        ,[Status] = '$($Object.Status)'
    WHERE UserID = '$($Object.UserID)'
    GO"

    if (Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryCheck)
    {
        # Update
        Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryUpdate

    }
    else {
        #Insert
        
        $Object | select-Object -Property UserID, User, Name, Role, TeamID, Status  | Write-SQLTableData -Serverinstance $StagingServer -DatabaseName $StagingDatabase -TableName $TableName -SchemaName "dbo"
    }

}

Function Reset-Deleted ($TableName)
{
    $QueryDeleted="USE [$StagingDatabase]
    GO

    UPDATE [dbo].[$TableName]
    SET [Status] = '0'
    GO"

    Invoke-Sqlcmd -ServerInstance $StagingServer -Database $StagingDatabase -Query $QueryDeleted



}


Write-au2matorLog -Type INFO -Text "Check For Module MicrosoftTeams"
if (Get-InstalledModule -Name "MicrosoftTeams") {
    Write-au2matorLog -Type INFO -Text "Module is installed"
    
}
else {
    Write-au2matorLog -Type INFO -Text "Module not found, try to install"
    Install-Module -Name MicrosoftTeams -Confirm:$false -Force
}

Write-au2matorLog -Type INFO -Text "Import TEAMS PS Module"
Import-Module MicrosoftTeams



Write-au2matorLog -Type INFO -Text "Check For Module sqlserver"
if (Get-InstalledModule -Name "sqlserver") {
    Write-au2matorLog -Type INFO -Text "Module is installed"
    
}
else {
    Write-au2matorLog -Type INFO -Text "Module not found, try to install"
    Install-Module -Name sqlserver -Confirm:$false -Force -AllowClobber
}

Write-au2matorLog -Type INFO -Text "Import SQL PS Module"
Import-Module sqlserver




# Make Sure Database exists
Write-au2matorLog -Type INFO -Text "Make sure Database exists"
try{
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
    $f_secTeamspasswd = ConvertTo-SecureString $TeamsAdminPW -AsPlainText -Force
    $f_myTeamscreds = New-Object System.Management.Automation.PSCredential ($TeamsAdminUser, $f_secTeamspasswd)
    
    Connect-MicrosoftTeams -Credential  $f_myTeamscreds
    
    try {
        $AllTeams=Get-Team
        Reset-Deleted -TableName "Teams-TeamsList"
        Reset-Deleted -TableName "Teams-ChannelList"
        Reset-Deleted -TableName "Teams-MemberList"
        foreach ($Team in $AllTeams)
        {
            Write-au2matorLog -Type INFO -Text "Working with Teams: $($Team.DisplayName))"
            Sync-Table_TeamsList -Object $Team 
        
        
        # Take care about Channels
            $AllChannels=Get-TeamChannel -GroupId $Team.GroupID
            foreach ($Channel in $AllChannels)
            {
                Write-au2matorLog -Type INFO -Text "Working with Channel: $($Channel.DisplayName))"
                Sync-Table_ChannelList -Object $Channel -TeamID $Team.GroupID
            }
        
            $AllMembers=Get-TeamUser -GroupId $Team.GroupId
            foreach ($Member in $AllMembers)
            {
                Write-au2matorLog -Type INFO -Text "Working with User: $($Member.User))"
                Sync-Table_MemberList -Object $Member -TeamID $Team.GroupID
            }
        }
    }
    catch {
        Write-au2matorLog -Type Error -Text "Error on sync: $Error"
    }
 
    
    Write-au2matorLog -Type INFO -Text "Sync finished"
}
catch {
    Write-au2matorLog -Type Error -Text "unable to connect to Teams Online: $Error"
}
