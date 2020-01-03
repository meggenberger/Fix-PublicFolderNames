<#
    .SYNOPSIS
    Script to prepare legacy public folder names for migration to modern public folders
   
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 0.1, 02.01.2020
    Ideas, comments and suggestions to me@devnull.ch 
 
    .LINK  
    https://github.com/meggenberger/Fix-PublicFolderNames
	
    .DESCRIPTION
	
    This script renames legacy public foldern on Exchange Server 2010 to replace backslash "\" and forward slash "/" by the pipe "|" character.
    This script trims all public folder names to remove any leading or trailing spaces.

    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
    - Exchange 2010 Management Shell
    Revision History 
    -------------------------------------------------------------------------------- 
    0.1     Initial version
    
    .PARAMETER PublicFolderServer
    Exchange server name hosting legacy public folders. If no server name is provided the server running the script is used
    
    .EXAMPLE
    Rename and trim public folders found on Server EXCHANGESERVER
    .\Fix-PublicFolderNames -PublicFolderServer EXCHANGESERVER
#>

[CmdletBinding()]
Param(
  [string] $PublicFolderServer 
)

#No Name specified. Using host name where the script is executed
if ($null -eq $PublicFolderServer -or "" -eq $PublicFolderServer) {
    Write-Debug("No server specified. Using local hostname {0}" -f $PublicFolderServer)
    $PublicFolderServer = $env:computername
    
}
function checkFolders($folders) {
    $percentDivider = $folders.Count / 100
    $counter = 0
    foreach($folder in $folders){
        $percent = [math]::Round($counter / $percentDivider)
        Write-Progress -Activity "Checking folders" -Status "$percent% complete:" -PercentComplete $percent
        Write-Debug("Checking folder {0}" -f $folder.Name)
        $needsUpdating = $false
        $newName = $folder.Name
        #check if there are leading or trailing spaces and tabs
        if($folder.Name -match '^[ \t]+' -or $folder.Name -match '[ \t]+$') {
            Write-Verbose("Found leading or trailing space in folder {0}" -f $folder.Name)
            $newName = $newName.Trim()
            $needsUpdating = $true
        }
        #check if there are forward or backwards slashes
        if($folder.Name -match '\/') {
            Write-Verbose("Found forward slash in name of folder {0}" -f $folder.Name)
            $newName = $newName.Replace('/','|')
            $needsUpdating = $true
        }
        if($folder.Name -match '\\') {
            Write-Verbose("Found back slash in name of folder {0}" -f $folder.Name)
            $newName = $newName.Replace('\','|')
            $needsUpdating = $true
        }
        #Only Update if it needs updating
        if($needsUpdating) {
            set-PublicFolder -Identity $folder.EntryId -Name $newName
            Write-Verbose("Folder {0} updated." -f $folder.Name)
        }
        $counter++
    }
} 

Write-Host("Getting all ipmSubtreeFolders on server {0}. This may take a while ..." -f $PublicFolderServer)
$ipmSubtreeFolders = Get-PublicFolder -Identity "\" -Server $PublicFolderServer -Recurse -ResultSize Unlimited
Write-Verbose("Found {0} folders on server {1} in ipmSubtree" -f $ipmSubtreeFolders.Count, $PublicFolderServer)

checkFolders($ipmSubtreeFolders)
   

    