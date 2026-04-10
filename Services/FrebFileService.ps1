#Requires -Version 5.1
<#
.SYNOPSIS
    File-discovery services for the FREB log viewer.

.DESCRIPTION
    Provides two functions:
      Get-FrebLogFiles  — enumerates FREB XML logs in a directory.
      Find-FrebXslFile  — locates the nearest freb.xsl stylesheet.
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    Returns all FREB XML log files in $Directory, newest first.

.PARAMETER Directory
    The folder to search.

.OUTPUTS
    [System.IO.FileInfo[]]  May be an empty array; never $null.
#>
function Get-FrebLogFiles {
    [CmdletBinding()]
    [OutputType([System.IO.FileInfo[]])]
    param(
        [Parameter(Mandatory)]
        [string] $Directory
    )

    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
        return @()
    }

    return @(
        Get-ChildItem -LiteralPath $Directory -Filter '*.xml' -File -ErrorAction SilentlyContinue |
            Sort-Object -Property LastWriteTime -Descending
    )
}

<#
.SYNOPSIS
    Finds the nearest freb.xsl file, starting from $XmlDirectory and walking upward.

.PARAMETER XmlDirectory
    The directory that contains the FREB XML files (search starts here).

.OUTPUTS
    [string]  Absolute path to the first freb.xsl found, or $null if none exists.
#>
function Find-FrebXslFile {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string] $XmlDirectory
    )

    $searchDir = $XmlDirectory
    while ($searchDir) {
        $candidate = Join-Path $searchDir 'freb.xsl'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return $candidate
        }

        $parent = Split-Path -Parent $searchDir
        # Stop when we reach the filesystem root (parent equals self)
        if (-not $parent -or $parent -eq $searchDir) { break }
        $searchDir = $parent
    }

    return $null
}

