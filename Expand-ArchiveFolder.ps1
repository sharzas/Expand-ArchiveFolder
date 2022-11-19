<#
    .SYNOPSIS
    Extract all rar, zip and 7z archives in a folder structure, to the specified destination
    directory.

    .DESCRIPTION
    Extract all rar, zip and 7z archives in a folder structure, to the specified destination
    directory.

    Will enumerate all archives, and use 7z.exe to extract them.
    
    The folder structure will be preserved in the destination directory, top folder being
    the name of the folder where the script is invoked from.

    E.g.: 
    Invoked from C:\Archives\Ledger with -Destination "D:\Extract" will extract create top
    folder D:\Extract\Ledger, and make sure all archives are extracted to the appropriate
    subdirectory in that directory tree.


    .PARAMETER Path
    Source root path where archives to be extracted are located. All subdirectories will
    we searched for archives.

    .PARAMETER Destination
    Root destination directory. The folder structure from -Path will be recreated here for
    any directory under -Path containing an archive to extract.
#>

[CmdletBinding()]

param (
    $Path = (Get-Location).Path,
    $Destination = "D:\Data\Downloads\Torrent\Extract"
)

if (!(Test-Path -Path $Path -PathType Container)) {
    throw ('Source Path "{0}" does not exist' -f $Path)

    try {
        # Get source path as DirectoryInfo object
        Get-Item -Path $SourcePath
    } catch {
        Write-Warning ('Unable to read directory "{0}"' -f $Path)
        Write-Host
        throw $_
    }
}

if (!(Test-Path -Path $Destination -PathType Container)) {
    throw ('Destination Path "{0}" does not exist' -f $Destination)
}


# Construct the full extract destination, by adding the root directory to the path
$ExtractPath = Join-Path -Path $Destination -ChildPath ($SourcePath | Split-Path -Leaf)

Write-Host ('Destination Extract Path: "{0}"' -f $ExtractPath)

# Enumerate all archives in folder structure - make sure to
# exclude everything but the main archive from multipart archives
$Archives = $SourcePath | Get-ChildItem -Recurse -File| where-object {($_.name-match '\.rar$|\.zip$|\.7z$') -and ($_.name -notmatch '\.(part1\d+?|part[2-9]*?)\.rar$' )}

# Create the list of archives and the constructed destination
$ExpandList = $Archives | Foreach-Object {
    [PSCustomObject]@{
        Archive = $_
        Destination = (Join-Path -Path $ExtractPath -ChildPath $_.Directory.FullName.Replace($SourcePath.Fullname, "")) -replace "\\$",""
    }
} #| select -first 1


# Loop through the list, and expand them
foreach ($Item in $ExpandList) {
    $Arguments = @(
        'x',
        '-o"""{0}"""' -f $Item.Destination,
        ""
        "-y",
        '"""{0}"""' -f $Item.Archive.Fullname
    )

    if (!(Test-Path $Item.Destination -PathType Container)) {
        # Destination directory does not exist - attempt to create it.
        Write-Verbose ('Create destination "{0}"' -f $Item.Destination)

        try {
            New-Item -Path $Item.Destination -ItemType Directory -ErrorAction Stop | Out-Null
        } catch {
            Write-Warning ('Error trying to create directory "{0}"' -f $Item.Destination)
            Write-Host
            throw $_
        }
    } else {
        Write-Verbose ('Destination already exist "{0}"' -f $Item.Destination)
    }

    Write-Host -ForegroundColor Cyan ('Extracting file : "{0}"' -f $Item.Archive.Fullname)
    Write-Host -ForegroundColor Cyan (' \-- Destination: "{0}"' -f $Item.Destination)
    Write-Verbose ('CommandLine: 7z {0}' -f ($Arguments -join " "))

    # Run command - redirect StdOut to NUL - we still want any errors to be seen.
    Start-Process 7z -ArgumentList $Arguments -NoNewWindow -Wait -RedirectStandardOutput "NUL"
}
