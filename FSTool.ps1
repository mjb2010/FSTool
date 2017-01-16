<#
.SYNOPSIS
FSTool: Folder Shortcut Tool v1.0

In the specified folder, and optionally all subfolders, finds all shortcuts to folders and replaces them with folder shell links.
Also reports on file & folder paths which exceed the 259-character limit of Windows APIs.

.PARAMETER foldername
The folder (directory) to start in. Defaults to '.', the current folder.

.DESCRIPTION
To use the script, launch a PowerShell instance, and in the console, type the name of the script, optionally followed by a
directory path (in quotes, if it contains spaces). The folder tree starting at that point, or the current folder if none is
given, will be recursively scanned for shortcut files pointing to folders. The script tells the user what they are, prompts
for confirmation, then converts the links to "folder shortcuts", a.k.a. "folder shell links".

A folder shell link is not an ordinary shell link (shortcut file). Instead, it is a special folder which, thanks to a hidden
desktop.ini file declaring it to be a shell link, behaves just like a symbolic link in Explorer. Other apps see it as what it
really is: a folder containing a hidden desktop.ini file, along with a regular shortcut file pointing to the target folder.

At some point this script may be enhanced, e.g. to reverse the process or to add support for converting to/from symlinks.

.NOTES
Author: Mike J. Brown <mike@skew.org>
License: CC0 1.0 <https://creativecommons.org/publicdomain/zero/1.0/>
Requires: Windows PowerShell 2.0
#>

Param([string]$foldername='.')
$error.clear()

Function GetShortcutTarget {
    # Given the path to a shortcut (.lnk) file, return the path of its target.
    # Beware: some shortcuts like RecentPlaces have empty targets.
    Param([string]$shortcutPath)
    Return (New-Object -COM WScript.Shell).CreateShortcut($shortcutPath).TargetPath
}

Function CreateTempDir {
    # In the system temp dir, make a temporary folder with a random name, and return the full path.
    $tmpDirPath = [System.IO.Path]::GetTempPath()
    $tmpDirPath = Join-Path -Path $tmpDirPath -ChildPath ([System.IO.Path]::GetRandomFileName())
    (New-Item -ItemType Directory -Path $tmpDirPath).FullName
}

# Get the folder .lnk files (as path strings) in the designated folder and subfolders,
# ignoring 'target.lnk' files, links with empty targets, and most error messages
# (e.g. for nonexistent targets); only keeping shortcuts which point to existing folders:
$folderLinkFilePaths = (
    Get-ChildItem -LiteralPath $foldername -Recurse -Include '*.lnk' -Exclude 'target.lnk' -ErrorAction SilentlyContinue -ErrorVariable linkDiscoveryErrs | where { 
        $targetPath = GetShortcutTarget($_.FullName)
        If ( -Not [string]::IsNullOrEmpty($targetPath) ) {
            (Get-Item -LiteralPath $targetPath -ErrorAction SilentlyContinue -ErrorVariable linkTargetErrs).PSIsContainer
        }
    }
).FullName

if ( $folderLinkFilePaths ) {
    Write-Host -ForegroundColor Green 'These shortcut files link to folders which exist:'
    $folderLinkFilePaths
} else {
    Write-Host -ForegroundColor Red 'Found no shortcut files, or none which link to folders which exist.'
}

$longFilePathFolders = ( $linkDiscoveryErrs | where { $_.Exception -is [System.IO.PathTooLongException] }).TargetObject
if ( $longFilePathFolders ) {
    $longPaths = ( $longFilePathFolders | foreach {
        $folderPath = $_
        $folderPathLength = $_.Length
        $subfolders = cmd /c "dir /b /ad ""$folderPath""" 2> $null
        ($subfolders | where { $_.Length + $folderPathLength -gt 259 }) | foreach { Join-Path -Path $folderPath -ChildPath $_ }
    } )
    if ( $longPaths ) {
        Write-Host ''
        Write-Host -ForegroundColor Red 'These unscanned subfolders have paths > 259 characters:'
        $longPaths | foreach {
            Write-Host -NoNewline -ForegroundColor Yellow 'Path length '
            Write-Host -NoNewline -ForegroundColor Red $_.Length
            Write-Host -NoNewline -ForegroundColor Yellow ': '
            Write-Host $_ 
        }
    }
    $longPaths = ( $longFilePathFolders | foreach {
        $folderPath = $_
        $folderPathLength = $_.Length
        $files = cmd /c "dir /b /a-d ""$folderPath""" 2> $null
        ($files | where { $_.Length + $folderPathLength -gt 259 }) | foreach { Join-Path -Path $folderPath -ChildPath $_ }
    } )
    if ( $longPaths ) {
        Write-Host ''
        Write-Host -ForegroundColor Red 'There are some files with paths > 259 characters:'
        $longPaths | foreach {
            Write-Host -NoNewline -ForegroundColor Yellow 'Path length '
            Write-Host -NoNewline -ForegroundColor Red $_.Length
            Write-Host -NoNewline -ForegroundColor Yellow ': '
            Write-Host $_ 
        }
    }
}

$missingTargets = ( $linkDiscoveryErrs | where { $_.Exception -is [System.Management.Automation.ItemNotFoundException] }).TargetObject
if ( $missingTargets ) {
    Write-Host ''
    Write-Host -ForegroundColor Red 'There are some shortcuts which point to nonexistent files or folders:'
    $missingTargets
}

$missingTargets = ( $linkDiscoveryErrs | where {  $_.Exception -is [System.Management.Automation.DriveNotFoundException] }).TargetObject
if ( $missingTargets ) {
    Write-Host ''
    Write-Host -ForegroundColor Red 'There are some shortcuts which point to files or folders on an offline or nonexistent drive:'
    $missingTargets
}

if ( $folderLinkFilePaths ) {
    $caption = "Choose"
    $message = "If you continue, shortcuts to existing folders will be replaced with shell link folders."
    $options = [System.Management.Automation.Host.ChoiceDescription[]] @("&Quit","&Continue")
    [int]$defaultOption = 1
    $answer = $host.ui.PromptForChoice($caption,$message,$options,$defaultOption)
    if ($answer) {
        $folderLinkFilePaths | foreach {
            Write-Host -NoNewline -ForegroundColor Yellow 'Processing shortcut: '
            Write-Host -NoNewline $_
            $shortcutFile = Get-Item -LiteralPath $_

            # make a temporary folder
            $newTmpFolderPath = CreateTempDir

            # in the temp folder, make a desktop.ini file with the necessary content and attributes
            $contentString = @'
[.ShellClassInfo]
CLSID2={0AFACED1-E828-11D1-9187-B532F1E9575D}
Flags=2
'@
            $desktopIniFile = New-Item -ItemType File -Path (Join-Path -Path $newTmpFolderPath -ChildPath 'desktop.ini') -Value $contentString
            $desktopIniFile.Attributes = 'Hidden,System'

            # in the temp folder, make a copy of the original shortcut, but name it target.lnk
            Copy-Item -LiteralPath $_ -Destination (Join-Path -Path $newTmpFolderPath -ChildPath 'target.lnk')

            # move the temp folder into the original shortcut's parent folder and rename it to the original shortcut's base name
            $linkFolderPath = Join-Path -Path $shortcutFile.Directory -ChildPath $shortcutFile.BaseName
            Move-Item -LiteralPath $newTmpFolderPath -Destination $linkFolderPath
            $linkFolder = Get-Item -LiteralPath $linkFolderPath

            # set the Read-Only attribute on the folder
            $linkFolder.Attributes = 'ReadOnly'

            # output a result message
            if ( $linkFolder ) {
                Write-Host -ForegroundColor Green ' [Replaced shortcut with shell link folder.]'
                # delete the shortcut
                $shortcutFile.Delete()
            } else {
                Write-Host -ForegroundColor Red ' [Something went wrong. Original shortcut remains intact.]'
            }
        }
    }
}