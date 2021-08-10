[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $ConversionConfigurationExportPath
,
    [Parameter()]
    [switch]
    $Exit
)

Function Validate-Dependencies {
    [CmdletBinding()]
    param ()

    # Validate assemblies
    if ( ($env:OS -imatch 'Windows') -and ! (Get-Item -Path $env:windir\assembly\GAC_MSIL\*onenote*) ) {
        "There are missing onenote assemblies. Please ensure the Desktop version of Onenote 2016 or above is installed." | Write-Warning
    }

    # Validate dependencies
    if (! (Get-Command -Name 'pandoc.exe') ) {
        throw "Could not locate pandoc.exe. Please ensure pandoc is installed."
    }
}

Function Get-DefaultConfiguration {
    [CmdletBinding()]
    param ()

    # The default configuration
    $config = [ordered]@{
        dryRun = @{
            description = @'
Whether to do a dry run
1: Convert - Default
2: Convert (dry run)
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        notesdestpath = @{
            description = @'
Specify folder path that will contain your resulting Notes structure - Default: c:\temp\notes
'@
            default = 'c:\temp\notes'
            value = 'c:\temp\notes'
            validateOptions = 'directoryexists'
        }
        targetNotebook = @{
            description = @'
Specify a notebook name to convert
'': Convert all notebooks - Default
'mynotebook': Convert specific notebook named 'mynotebook'
'@
            default = ''
            value = ''
        }
        usedocx = @{
            description = @'
Whether to create new word .docx or reuse existing ones
1: Always create new .docx files - Default
2: Use existing .docx files (90% faster)
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        keepdocx = @{
            description = @'
Whether to discard word .docx after conversion
1: Discard intermediate .docx files - Default
2: Keep .docx files
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        prefixFolders = @{
            description = @'
Whether to use prefix vs subfolders
1: Create folders for subpages (e.g. Page\Subpage.md) - Default
2: Add prefixes for subpages (e.g. Page_Subpage.md)
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        medialocation = @{
            description = @'
Whether to store media in single or multiple folders
1: Images stored in single 'media' folder at Notebook-level - Default
2: Separate 'media' folder for each folder in the hierarchy
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        conversion = @{
            description = @'
Specify conversion type
1: markdown (Pandoc) - Default
2: commonmark (CommonMark Markdown)
3: gfm (GitHub-Flavored Markdown)
4: markdown_mmd (MultiMarkdown)
5: markdown_phpextra (PHP Markdown Extra)
6: markdown_strict (original unextended Markdown)
'@
            default = 1
            value = 1
            validateRange = 1,6
        }
        headerTimestampEnabled = @{
            description = @'
Whether to include page timestamp and separator at top of document
1: Include - Default
2: Don't include
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        keepspaces = @{
            description = @'
Whether to clear double spaces between bullets
1: Clear double spaces in bullets - Default
2: Keep double spaces
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        keepescape = @{
            description = @'
Whether to clear escape symbols from md files
1: Clear '\' symbol escape character from files - Default
2: Keep '\' symbol escape
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
    }

    $config
}

Function New-ConfigurationFile {
    [CmdletBinding()]
    param ()

    # Generate a configuration file config.example.ps1
    @'
#
# Note: This config file is for those who are lazy to type in configuration everytime you run ./ConvertOneNote2MarkDown-v2.ps1
#
# Steps:
#   1) Rename this file to config.ps1. Ensure it is in the same folder as the ConvertOneNote2MarkDown-v2.ps1 script
#   2) Configure the options below to your liking
#   3) Run the main script: ./ConvertOneNote2MarkDown-v2.ps1. Sit back while the script starts converting immediately.
'@ | Out-File "$PSScriptRoot/config.example.ps1" -Encoding utf8

    $defaultConfig = Get-DefaultConfiguration
    foreach ($key in $defaultConfig.Keys) {
        # Add a '#' in front of each line of the option description
        $defaultConfig[$key]['description'].Trim() -replace "^|`n", "`n# " | Out-File "$PSScriptRoot/config.example.ps1" -Encoding utf8 -Append

        # Write the variable
        if ( $defaultConfig[$key]['default'] -is [string]) {
            "`$$key = '$( $defaultConfig[$key]['default'] )'" | Out-File "$PSScriptRoot/config.example.ps1" -Encoding utf8 -Append
        }else {
            "`$$key = $( $defaultConfig[$key]['default'] )" | Out-File "$PSScriptRoot/config.example.ps1" -Encoding utf8 -Append
        }
    }
}

Function Compile-Configuration {
    [CmdletBinding()]
    param ()

    # Get a default configuration
    $config = Get-DefaultConfiguration

    # Override configuration
    if (Test-Path $PSScriptRoot/config.ps1) {
        # Get override configuration from config file ./config.ps1
        & {
            $scriptblock = [scriptblock]::Create( (Get-Content $PSScriptRoot/config.ps1 -Raw) )
            . $scriptblock
            foreach ($key in @($config.Keys)) {
                $config[$key]['value'] = Get-Variable -Name $key -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Value
                # Trim string
                if ($config[$key]['value'] -is [string]) {
                    $config[$key]['value'] = $config[$key]['value'].Trim()
                }
                # Remove trailing slash(es) for paths
                if ($key -match 'path' -and $config[$key]['value'] -match '[/\\]') {
                    $config[$key]['value'] = $config[$key]['value'].TrimEnd('/').TrimEnd('\')
                }
            }
        }
    }else {
        # Get override configuration from interactive prompts
        foreach ($key in $config.Keys) {
            "" | Write-Host -ForegroundColor Cyan
            $config[$key]['description'] | Write-Host -ForegroundColor Cyan
            # E.g. 'string', 'int'
            $typeName = [Microsoft.PowerShell.ToStringCodeMethods]::Type($config[$key]['default'].GetType())
            # Keep prompting until we get a answer of castable type
            do {
                # Cast the input as a type. E.g. Read-Host -Prompt 'Entry' -as [int]
                $config[$key]['value'] = Invoke-Expression -Command "(Read-Host -Prompt 'Entry') -as [$typeName]"
            }while ($null -eq $config[$key]['value'])
            # Fallback on default value if the input is empty string
            if ($config[$key]['value'] -is [string] -and $config[$key]['value'] -eq '') {
                $config[$key]['value'] = $config[$key]['default']
            }
            # Fallback on default value if the input is empty integer (0)
            if ($config[$key]['value'] -is [int] -and $config[$key]['value'] -eq 0) {
                $config[$key]['value'] = $config[$key]['default']
            }
        }
    }

    $config
}

Function Validate-Configuration {
    [CmdletBinding(DefaultParameterSetName='default')]
    param (
        [Parameter(ParameterSetName='default',Position=0)]
        [object]
        $Config
    ,
        [Parameter(ParameterSetName='pipeline',ValueFromPipeline)]
        [object]
        $InputObject
    )
    process {
        if ($InputObject) {
            $Config = $InputObject
        }
        if ($null -eq $Config) {
            throw "No input parameters specified."
        }

        # Validate a given configuration against a prototype configuration
        $defaultConfig = Get-DefaultConfiguration
        foreach ($key in $defaultConfig.Keys) {
            if (! $Config.Contains($key)) {
                throw "Missing configuration option '$key'"
            }
            if ($defaultConfig[$key]['default'].GetType().FullName -ne $Config[$key]['value'].GetType().FullName) {
                throw "Invalid configuration option '$key'. Expected a value of type $( $defaultConfig[$key]['default'].GetType().FullName ), but value was of type $( $config[$key]['value'].GetType().FullName )"
            }
            if ($defaultConfig[$key].Contains('validateOptions')) {
                if ($defaultConfig[$key]['validateOptions'] -contains 'directoryexists') {
                    if ( ! $config[$key]['value'] -or ! (Test-Path $config[$key]['value'] -PathType Container -ErrorAction SilentlyContinue) ) {
                        throw "Invalid configuration option '$key'. The directory '$( $config[$key]['value'] )' does not exist, or is a file"
                    }
                }
            }
            if ($defaultConfig[$key].Contains('validateRange')) {
                if ($Config[$key]['value'] -lt $defaultConfig[$key]['validateRange'][0] -or $Config[$key]['value'] -gt $defaultConfig[$key]['validateRange'][1]) {
                    throw "Invalid configuration option '$key'. The value must be between $( $defaultConfig[$key]['validateRange'][0] ) and $( $defaultConfig[$key]['validateRange'][1] )"
                }
            }
        }

        # Warn of unknown configuration options
        foreach ($key in $config.Keys) {
            if (! $defaultConfig.Contains($key)) {
                "Unknown configuration option '$key'" | Write-Warning
            }
        }

        $Config
    }
}

Function Print-Configuration {
    [CmdletBinding(DefaultParameterSetName='default')]
    param (
        [Parameter(ParameterSetName='default',Position=0)]
        [object]
        $Config
    ,
        [Parameter(ParameterSetName='pipeline',ValueFromPipeline)]
        [object]
        $InputObject
    )
    process {
        if ($InputObject) {
            $Config = $InputObject
        }
        if ($null -eq $Config) {
            throw "No input parameters specified."
        }

        foreach ($key in $Config.Keys) {
            "$( $key ): $( $Config[$key]['value'] )" | Write-Host -ForegroundColor DarkGray
        }
    }
}

Function Remove-InvalidFileNameChars {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline = $true)]
        [AllowEmptyString()]
        [string]$Name
    ,
        [switch]$KeepPathSpaces
    )

    # Remove boundary whitespaces. So we don't get trailing dashes
    $Name = $Name.Trim()

    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    $newName = $newName -replace "\[", "("
    $newName = $newName -replace "\]", ")"
    $newName =  if ($KeepPathSpaces) {
                    $newName -replace "\s", " "
                } else {
                    $newName -replace "\s", "-"
                }
    $newName = $newName.Substring(0, $(@{$true = 130; $false = $newName.length }[$newName.length -gt 150]))
    return $newName
}

Function Remove-InvalidFileNameCharsInsertedFiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]$Name
    ,
        [string]$Replacement = ""
    ,
        [string]$SpecialChars = "#$%^*[]'<>!@{};"
    ,
        [switch]$KeepPathSpaces
    )

    # Remove boundary whitespaces. So we don't get trailing dashes
    $Name = $Name.Trim()

    $rePattern = ($SpecialChars.ToCharArray() | ForEach-Object { [regex]::Escape($_) }) -join "|"

    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    $newName = $newName -replace $rePattern, ""
    $newName =  if ($KeepPathSpaces) {
                    $newName -replace "\s", " "
                } else {
                    $newName -replace "\s", "-"
                }
    return $newName
}

Function New-OneNoteConnection {
    [CmdletBinding()]
    param ()

    # Create a OneNote connection. See: See: https://docs.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote
    if ($PSVersionTable.PSVersion.Major -le 5) {
        if ($OneNote = New-Object -ComObject OneNote.Application) {
            $OneNote
        }else {
            throw "Failed to make connection to OneNote."
        }
    }else {
        # Works between powershell 6.0 and 7.0, but not >= 7.1
        if (Add-Type -Path $env:windir\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll -PassThru) {
            $OneNote = [Microsoft.Office.Interop.OneNote.ApplicationClass]::new()
            $OneNote
        }else {
            throw "Failed to make connection to OneNote."
        }
    }
}

Function Remove-OneNoteConnection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    )

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNoteConnection) | Out-Null
}

Function Get-OneNoteHierarchy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $OneNoteConnection
    )

    # Open OneNote hierarchy
    [xml]$hierarchy = ""
    $OneNoteConnection.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$hierarchy)

    $hierarchy
}

Function Get-OneNotePageContent {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $OneNoteConnection
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $PageId
    )

    # Get page's xml content
    [xml]$page = ""
    $OneNoteConnection.GetPageContent($PageId, [ref]$page, 7)

    $page
}

Function Publish-OneNotePageToDocx {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $OneNoteConnection
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $PageId
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Destination
    )

    $OneNoteConnection.Publish($PageId, $Destination, "pfWord", "")
}

Function New-SectionGroupConversionConfig {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    ,
        # The desired directory to store any converted Page(s) found in this Section Group's Section(s)
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $NotesDestination
    ,
        [Parameter(Mandatory)]
        [object]
        $Config
    ,
        # Section Group XML object(s)
        [Parameter(Mandatory)]
        [array]
        $SectionGroups
    ,
        [Parameter(Mandatory)]
        [int]
        $LevelsFromRoot
    ,
        [Parameter()]
        [switch]
        $AsArray
    )

    $sectionGroupConversionConfig = [System.Collections.ArrayList]@()

    # Build an object representing the conversion of a Section Group (treat a Notebook as a Section Group, it is no different)
    foreach ($sectionGroup in $SectionGroups) {
        # Skip over Section Groups in recycle bin
        if ((Get-Member -InputObject $sectionGroup -Name 'isRecycleBin') -and $sectionGroup.isRecycleBin -eq 'true') {
            continue
        }

        $cfg = [ordered]@{}

        if ($LevelsFromRoot -eq 0) {
            "`nBuilding conversion configuration for $( $sectionGroup.name ) [Notebook]" | Write-Host -ForegroundColor DarkGreen
        }else {
            "`n$( '#' * ($LevelsFromRoot) ) Building conversion configuration for $( $sectionGroup.name ) [Section Group]" | Write-Host -ForegroundColor DarkGray
        }

        # Build this Section Group
        $cfg = [ordered]@{}
        $cfg['object'] = $sectionGroup # Keep a reference to the SectionGroup object
        $cfg['kind'] = 'SectionGroup'
        $cfg['nameCompat'] = $sectionGroup.name | Remove-InvalidFileNameChars
        $cfg['levelsFromRoot'] = $LevelsFromRoot
        $cfg['uri'] = $sectionGroup.path # E.g. https://d.docs.live.net/0123456789abcdef/Skydrive Notebooks/mynotebook/mysectiongroup
        $cfg['notesDirectory'] = [io.path]::combine( $NotesDestination.TrimEnd('/').TrimEnd('\').Replace('\', [io.path]::DirectorySeparatorChar), $cfg['nameCompat'] )
        $cfg['notesBaseDirectory'] = & {
            # E.g. 'c:\temp\notes\mynotebook\mysectiongroup'
            # E.g. levelsFromRoot: 1
            $split = $cfg['notesDirectory'].Split( [io.path]::DirectorySeparatorChar )
            # E.g. 5
            $totalLevels = $split.Count
            # E.g. 0..(5-1-1) -> 'c:\temp\notes\mynotebook'
            $split[0..($totalLevels - $cfg['levelsFromRoot'] - 1)] -join [io.path]::DirectorySeparatorChar
        }
        $cfg['pathFromRoot'] = $cfg['notesDirectory'].Replace($cfg['notesBaseDirectory'], '').Trim([io.path]::DirectorySeparatorChar)
        $cfg['pathFromRootCompat'] = $cfg['pathFromRoot'] | Remove-InvalidFileNameChars
        $cfg['notesDocxDirectory'] = [io.path]::combine( $cfg['notesBaseDirectory'], 'docx' )
        $cfg['directoriesToCreate'] = @()

        # Build this Section Group's sections
        $cfg['sections'] = [System.Collections.ArrayList]@()
        foreach ($section in $sectionGroup.Section) {
            "$( '#' * ($LevelsFromRoot + 1) ) Building conversion configuration for $( $section.name ) [Section]" | Write-Host -ForegroundColor DarkGray

            $sectionCfg = [ordered]@{}
            $sectionCfg['notesBaseDirectory'] = $cfg['notesBaseDirectory']
            $sectionCfg['notesDirectory'] = $cfg['notesDirectory']
            $sectionCfg['sectionGroupUri'] = $cfg['uri'] # Keep a reference to mt Section Group Configuration object's uri
            $sectionCfg['sectionGroupName'] = $cfg['object'].name
            $sectionCfg['object'] = $section # Keep a reference to the Section object
            $sectionCfg['kind'] = 'Section'
            $sectionCfg['nameCompat'] = $section.name | Remove-InvalidFileNameChars
            $sectionCfg['levelsFromRoot'] = $cfg['levelsFromRoot'] + 1
            $sectionCfg['pathFromRoot'] = "$( $cfg['pathFromRoot'] )$( [io.path]::DirectorySeparatorChar )$( $sectionCfg['nameCompat'] )".Trim([io.path]::DirectorySeparatorChar)
            $sectionCfg['pathFromRootCompat'] = $sectionCfg['pathFromRoot'] | Remove-InvalidFileNameChars
            $sectionCfg['uri'] = $section.path # E.g. https://d.docs.live.net/0123456789abcdef/Skydrive Notebooks/mynotebook/mysectiongroup/mysection
            $sectionCfg['lastModifiedTime'] = [Datetime]::ParseExact($section.lastModifiedTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
            $sectionCfg['pages'] = [System.Collections.ArrayList]@()

            # Build Section's pages
            foreach ($page in $section.Page) {
                "$( '#' * ($LevelsFromRoot + 2) ) Building conversion configuration for $( $page.name ) [Page]" | Write-Host -ForegroundColor DarkGray

                $previousPage = if ($sectionCfg['pages'].Count -gt 0) { $sectionCfg['pages'][$sectionCfg['pages'].Count - 1] } else { $null }
                $pageCfg = [ordered]@{}
                $pageCfg['notesBaseDirectory'] = $cfg['notesBaseDirectory']
                $pageCfg['notesDirectory'] = $cfg['notesDirectory']
                $pageCfg['sectionGroupUri'] = $cfg['uri'] # Keep a reference to mt Section Group Configuration object's uri
                $pageCfg['sectionGroupName'] = $cfg['object'].name
                $pageCfg['sectionUri'] = $sectionCfg['uri'] # Keep a reference to my Section Configuration object's uri
                $pageCfg['sectionName'] = $sectionCfg['object'].name
                $pageCfg['object'] = $page # Keep a reference to my Page object
                $pageCfg['kind'] = 'Page'
                $pageCfg['nameCompat'] = $page.name | Remove-InvalidFileNameChars
                $pageCfg['levelsFromRoot'] = $sectionCfg['levelsFromRoot']
                $pageCfg['pathFromRoot'] = "$( $sectionCfg['pathFromRoot'] )$( [io.path]::DirectorySeparatorChar )$( $pageCfg['nameCompat'] )"
                $pageCfg['pathFromRootCompat'] = $pageCfg['pathFromRoot'] | Remove-InvalidFileNameChars
                $pageCfg['uri'] = "$( $sectionCfg['object'].path )/$( $page.name )" # There's no $page.path property, so we generate one. E.g. https://d.docs.live.net/0123456789abcdef/Skydrive Notebooks/mynotebook/mysectiongroup/mysection/mypage
                $pageCfg['dateTime'] = [Datetime]::ParseExact($page.dateTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                $pageCfg['lastModifiedTime'] = [Datetime]::ParseExact($page.lastModifiedTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                $pageCfg['pageLevel'] = $page.pageLevel -as [int]
                $pageCfg['converter'] = switch ($config['conversion']['value']) {
                    1 { 'markdown' }
                    2 { 'commonmark' }
                    3 { 'gfm' }
                    4 { 'markdown_mmd' }
                    5 { 'markdown_phpextra' }
                    6 { 'markdown_strict' }
                    default { 'markdown' }
                }
                $pageCfg['pagePrefix'] = & {
                    if ($pageCfg['pageLevel'] -eq 1) {
                        ''
                    }else {
                        if ($previousPage) {
                            if ($previousPage['pageLevel'] -lt $pageCfg['pageLevel']) {
                                # 1 -> 2, 1 -> 3, or 2 -> 3
                                "$( $previousPage['filePathRel'] )$( [io.path]::DirectorySeparatorChar )"
                            }elseif ($previousPage['pageLevel'] -eq $pageCfg['pageLevel']) {
                                # 2 -> 2, or 3 -> 3
                                "$( Split-Path $previousPage['filePathRel'] -Parent )$( [io.path]::DirectorySeparatorChar )"
                            }else {
                                # 3 -> 2
                                $previousLevel1Page = @(
                                    $sectionCfg['pages'] | Where-Object { $_['pageLevel'] }
                                )
                                if ($previousLevel1Page.Count -gt 0) {
                                    "$( Split-Path (Split-Path $previousPage['filePathRel'] -Parent) -Parent )$( [io.path]::DirectorySeparatorChar )"
                                }else {
                                    ''
                                }
                            }
                        }else {
                            ''
                        }
                    }
                }
                $pageCfg['filePathRel'] = & {
                    $filePathRel = "$( $pageCfg['pagePrefix'] )$( $pageCfg['nameCompat'] )"

                    # in case multiple pages with the same name exist in a section, postfix the filename
                    $recurrence = 0
                    foreach ($p in $sectionCfg['pages']) {
                        if ($p['pagePrefix'] -eq $pageCfg['pagePrefix'] -and $p['pathFromRoot'] -eq $pageCfg['pathFromRoot']) {
                            $recurrence++
                        }
                    }
                    if ($recurrence -gt 0) {
                        $filePathRel = "$filePathRel-$recurrence"
                    }
                    $filePathRel
                }
                $pageCfg['filePathRelUnderscore'] = $pageCfg['filePathRel'].Replace( [io.path]::DirectorySeparatorChar, '_' )
                $pageCfg['fullfilepathwithoutextension'] = if ($config['prefixFolders']['value'] -eq 2) {
                    [io.path]::combine( $cfg['notesDirectory'], $sectionCfg['nameCompat'], $pageCfg['filePathRelUnderscore'] )
                }else {
                    [io.path]::combine( $cfg['notesDirectory'], $sectionCfg['nameCompat'], $pageCfg['filePathRel'] )
                }
                $pageCfg['fullexportdirpath'] = Split-Path $pageCfg['fullfilepathwithoutextension'] -Parent
                $pageCfg['mdFileName'] = Split-Path $pageCfg['fullfilepathwithoutextension'] -Leaf
                $pageCfg['levelsPrefix'] = if ($config['medialocation']['value'] -eq 2) {
                    ''
                }else {
                    if ($config['prefixFolders']['value'] -eq 2) {
                        "$( '../' * ($pageCfg['levelsFromRoot'] + 1 - 1) )"
                    }else {
                        "$( '../' * ($pageCfg['levelsFromRoot'] + $pageCfg['pageLevel'] - 1) )"
                    }
                }
                $pageCfg['mediaParentPath'] = if ($config['medialocation']['value'] -eq 2) {
                    $pageCfg['fullexportdirpath']
                }else {
                    $cfg['notesBaseDirectory']
                }
                $pageCfg['mediaPath'] = [io.path]::combine( $pageCfg['mediaParentPath'], 'media' )
                $pageCfg['mediaParentPath'] = Split-Path $pageCfg['mediaPath'] -Parent
                $pageCfg['mediaPathPandoc'] = $pageCfg['mediaPath'].Replace( [io.path]::DirectorySeparatorChar, '/' ) # Pandoc outputs paths in markdown with with front slahes after the supplied <mediaPath>, e.g. '<mediaPath>/media/image.png'. So let's use a front-slashed supplied mediaPath
                $pageCfg['mediaParentPathPandoc'] = (Split-Path $pageCfg['mediaPathPandoc'] -Parent).Replace( [io.path]::DirectorySeparatorChar, '/' ) # Pandoc outputs paths in markdown with with front slahes after the supplied <mediaPath>, e.g. '<mediaPath>/media/image.png'. So let's use a front-slashed supplied mediaPath
                $pageCfg['fullexportpath'] = [io.path]::combine( $cfg['notesDocxDirectory'], "$( $pageCfg['pathFromRootCompat'] ).docx" )
                $pageCfg['insertedAttachments'] = @(
                    & {
                        $pagexml = Get-OneNotePageContent -OneNoteConnection $OneNoteConnection -PageId $pageCfg['object'].ID

                        # Get any attachment(s) found in pages
                        if (Get-Member -InputObject $pagexml.Page -Name 'Outline') {
                            $insertedFiles = $pagexml.Page.Outline.OEChildren.OE | Where-Object { $null -ne $_ -and (Get-Member -InputObject $_ -Name 'InsertedFile') } | ForEach-Object { $_.InsertedFile }
                            foreach ($i in $insertedFiles) {
                                $attachmentCfg = [ordered]@{}
                                $attachmentCfg['object'] =  $i
                                $attachmentCfg['nameCompat'] =  $i.preferredName | Remove-InvalidFileNameCharsInsertedFiles
                                $attachmentCfg['markdownFileName'] =  $attachmentCfg['nameCompat'].Replace("$", "\$").Replace("^", "\^").Replace("'", "\'")
                                $attachmentCfg['source'] =  $i.pathCache
                                $attachmentCfg['destination'] =  [io.path]::combine( $pageCfg['mediaPath'], $attachmentCfg['nameCompat'] )

                                $attachmentCfg
                            }
                        }
                    }
                )
                $pageCfg['mutations'] = @(
                    # Markdown mutations

                    foreach ($attachmentCfg in $pageCfg['insertedAttachments']) {
                        @{
                            description = 'Change inserted attachment(s) filename references'
                            replacements = @(
                                @{
                                    searchRegex = [regex]::Escape( $attachmentCfg['object'].preferredName )
                                    replacement = "[$( $attachmentCfg['markdownFileName'] )]($( $pageCfg['mediaPathPandoc'] )/$( $attachmentCfg['markdownFileName'] ))"
                                }
                            )
                        }
                    }
                    @{
                        description = 'Replace media (e.g. images, attachments) absolute paths with relative paths'
                        replacements = @(
                            @{
                                # E.g. 'C:/temp/notes/mynotebook/media/somepage-image1-timestamp.jpg' -> '../media/somepage-image1-timestamp.jpg'
                                searchRegex = [regex]::Escape("$( $pageCfg['mediaParentPathPandoc'] )/") # Add a trailing front slash
                                replacement = $pageCfg['levelsPrefix']
                            }
                        )
                    }
                    @{
                        description = 'Add heading'
                        replacements = @(
                            @{
                                searchRegex = '^[^\r\n]*'
                                replacement = & {
                                    $heading = "# $( $pageCfg['object'].name )"
                                    if ($config['headerTimestampEnabled']['value'] -eq 1) {
                                        $heading += $pageCfg['dateTime'].ToString("`n`nyyyy-MM-dd HH:mm:ss")
                                        $heading += "`r`n`r`n---`r`n"
                                    }
                                    $heading
                                }
                            }
                        )
                    }
                    if ($config['keepspaces']['value'] -eq 1 ) {
                        @{
                            description = 'Clear double spaces from bullets and non-breaking spaces spaces from blank lines'
                            replacements = @(
                                @{
                                    searchRegex = [regex]::Escape([char]0x00A0)
                                    replacement = ''
                                }
                                @{
                                    searchRegex = '\r?\n\r?\n- '
                                    replacement = "`r`n- "
                                }
                            )
                        }
                    }
                    if ($config['keepescape']['value'] -eq 1) {
                        @{
                            description = 'Clear backslash escape symbols'
                            replacements = @(
                                @{
                                    searchRegex = [regex]::Escape('\')
                                    replacement = ''
                                }
                            )
                        }
                    }
                )
                $pageCfg['directoriesToCreate'] = @(
                    # The directories to be created
                    @(
                        $cfg['notesDocxDirectory']
                        $cfg['notesDirectory']
                        $pageCfg['fullexportdirpath']
                        $pageCfg['mediaPath']
                    ) | Select-Object -Unique
                )
                $pageCfg['directorySeparatorChar'] = [io.path]::DirectorySeparatorChar

                # Populate the pages array (needed even when -AsArray switch is not on, because we need this section's pages' state to know whether there are duplicate page names)
                $sectionCfg['pages'].Add( $pageCfg ) > $null

                if (!$AsArray) {
                    # Send the configuration immediately down the pipeline
                    $pageCfg
                }
            }

            # Populate the sections array
            if ($AsArray) {
                $cfg['sections'].Add( $sectionCfg ) > $null
            }
        }

        # Build this Section Group's Section Groups
        if ((Get-Member -InputObject $sectionGroup -Name 'SectionGroup')) {
            if ($AsArray) {
                $cfg['sectionGroups'] = New-SectionGroupConversionConfig -OneNoteConnection $OneNoteConnection -NotesDestination $cfg['notesDirectory'] -Config $Config -SectionGroups $sectionGroup.SectionGroup -LevelsFromRoot ($LevelsFromRoot + 1) -AsArray:$AsArray
            }else {
                # Send the configuration immediately down the pipeline
                New-SectionGroupConversionConfig -OneNoteConnection $OneNoteConnection -NotesDestination $cfg['notesDirectory'] -Config $Config -SectionGroups $sectionGroup.SectionGroup -LevelsFromRoot ($LevelsFromRoot + 1)
            }
        }

        # Populate the conversion config
        if ($AsArray) {
            $sectionGroupConversionConfig.Add( $cfg ) > $null
        }
    }

    # Return the final conversion config
    if ($AsArray) {
        ,$sectionGroupConversionConfig # This syntax is needed to send an array down the pipeline without it being unwrapped. (It works by wrapping it in an array with a null sibling)
    }
}

Function Convert-OneNotePage {
    [CmdletBinding(DefaultParameterSetName='default')]
    param (
        # Onenote connection object
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    ,
        # ConvertOneNote2MarkDown configuration object
        [Parameter(Mandatory)]
        [object]
        $Config
    ,
        # Conversion object
        [Parameter(Mandatory,ParameterSetName='default')]
        [ValidateNotNullOrEmpty()]
        [object]
        $ConversionConfig
    ,
        [Parameter(Mandatory,ParameterSetName='pipeline',ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [object]
        $InputObject
    )

    process {
        if ($InputObject) {
            $ConversionConfig = $InputObject
        }
        if ($null -eq $ConversionConfig) {
            throw "No config specified."
        }

        try {
            $pageCfg = $ConversionConfig

            "$( '#' * ($pageCfg['levelsFromRoot'] + $pageCfg['pageLevel']) ) $( $pageCfg['object'].name ) [$( $pageCfg['kind'] )]" | Write-Host
            "Uri: $( $pageCfg['uri'] )" | Write-Verbose

            # Create directories
            foreach ($d in $pageCfg['directoriesToCreate']) {
                try {
                    "Directory: $( $d )" | Write-Verbose
                    if ($config['dryRun']['value'] -eq 1) {
                        $item = New-Item -Path $d -ItemType Directory -Force -ErrorAction Stop
                    }
                }catch {
                    throw "Failed to create directory '$d': $( $_.Exception.Message )"
                }
            }

            if ($config['usedocx']['value'] -eq 1) {
                # Remove any existing docx files, don't proceed if it fails
                try {
                    "Removing existing docx file: $( $pageCfg['fullexportpath'] )" | Write-Verbose
                    if ($config['dryRun']['value'] -eq 1) {
                        if (Test-Path $pageCfg['fullexportpath']) {
                            Remove-Item -path $pageCfg['fullexportpath'] -Force -ErrorAction Stop
                        }
                    }
                }catch {
                    throw "Error removing intermediary docx file $( $pageCfg['fullexportpath'] ): $( $_.Exception.Message )"
                }
            }

            # Publish OneNote page to Word, don't proceed if it fails
            if (! (Test-Path $pageCfg['fullexportpath']) ) {
                try {
                    "Publishing new docx file: $( $pageCfg['fullexportpath'] )" | Write-Verbose
                    if ($config['dryRun']['value'] -eq 1) {
                        Publish-OneNotePageToDocx -OneNoteConnection $OneNoteConnection -PageId $pageCfg['object'].ID -Destination $pageCfg['fullexportpath']
                    }
                }catch {
                    throw "Error while publishing page to docx file $( $pageCfg['object'].name ): $( $_.Exception.Message )"
                }
            }else {
                "Existing docx file: $( $pageCfg['fullexportpath'] )" | Write-Verbose
            }

            # https://gist.github.com/heardk/ded40b72056cee33abb18f3724e0a580
            # Convert .docx to .md, don't proceed if it fails
            try {
                "Converting docx file to markdown file: $( $pageCfg['fullfilepathwithoutextension'] ).md" | Write-Verbose
                if ($config['dryRun']['value'] -eq 1) {
                    $process = Start-Process -PassThru -NoNewWindow -Wait -FilePath pandoc.exe -ArgumentList @( '-f', 'docx', '-t', "$( $pageCfg['converter'] )-simple_tables-multiline_tables-grid_tables+pipe_tables", '-i', $pageCfg['fullexportpath'], '-o', "$( $pageCfg['fullfilepathwithoutextension'] ).md", '--wrap=none', '--markdown-headings=atx', "--extract-media=$( $pageCfg['mediaParentPathPandoc'] )" ) # extracts into ./media of the supplied folder
                    if ($process.ExitCode) {
                        throw 'pandoc failed to convert'
                    }
                }
            }catch {
                throw "Error while converting docx file $( $pageCfg['fullexportpath'] ) to markdown file $( $pageCfg['fullfilepathwithoutextension'] ).md: $( $_.Exception.Message )"
            }

            # Cleanup Word files
            if ($config['keepdocx']['value'] -eq 1) {
                try {
                    "Removing existing docx file: $( $pageCfg['fullexportpath'] )" | Write-Verbose
                    if ($config['dryRun']['value'] -eq 1) {
                        if (Test-Path $pageCfg['fullexportpath']) {
                            Remove-Item -path $pageCfg['fullexportpath'] -Force -ErrorAction Stop
                        }
                    }
                }catch {
                    Write-Error "Error removing intermediary docx file $( $pageCfg['fullexportpath'] ): $( $_.Exception.Message )"
                }
            }

            # Save any attachments
            foreach ($attachmentCfg in $pageCfg['insertedAttachments']) {
                try {
                    "Saving inserted attachment: $( $attachmentCfg['destination'] )" | Write-Verbose
                    if ($config['dryRun']['value'] -eq 1) {
                        Copy-Item -Path $attachmentCfg['source'] -Destination $attachmentCfg['destination'] -Force -ErrorAction Stop
                    }
                }catch {
                    Write-Error "Error while saving attachment from $( $attachmentCfg['source'] ) to $( $attachmentCfg['destination'] ): $( $_.Exception.Message )"
                }
            }

            # Rename images to have unique names - NoteName-Image#-HHmmssff.xyz
            $timeStamp = (Get-Date -Format HHmmssff).ToString()
            $timeStamp = $timeStamp.replace(':', '')
            $images = Get-ChildItem -Path "$( $pageCfg['mediaPath'] )" -Include "*.png", "*.gif", "*.jpg", "*.jpeg" -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name.SubString(0, 5) -match "image" }
            foreach ($image in $images) {
                $newimageName = "$($( $pageCfg['mdFileName'] ).SubString(0,[math]::min(30,$( $pageCfg['mdFileName'] ).length)))-$($image.BaseName)-$($timeStamp)$($image.Extension)"
                # Rename Image
                try {
                    "Renaming image: $( $image.FullName ) to $( $newimageName )" | Write-Verbose
                    if ($config['dryRun']['value'] -eq 1) {
                        $item = Rename-Item -Path "$( $image.FullName )" -NewName $newimageName -ErrorAction Stop -PassThru
                    }
                }catch {
                    Write-Error "Error while renaming image $( $image.FullName ) to $( $item.FullName ): $( $_.Exception.Message )"
                }
                # Change MD file Image filename References
                try {
                    "Mutation of markdown: Rename image references to unique name" | Write-Verbose
                    if ($config['dryRun']['value'] -eq 1) {
                        $content = Get-Content -Path "$( $pageCfg['fullfilepathwithoutextension'] ).md" -Raw -ErrorAction Stop # Get-Content -ErrorAction Stop can produce random "Cannot find path 'xxx' because it does not exist"
                        $content = $content.Replace("$($image.Name)", "$($newimageName)")
                        Set-Content -Path "$( $pageCfg['fullfilepathwithoutextension'] ).md" -Value $content -ErrorAction Stop
                    }
                }catch {
                    Write-Error "Error while renaming image file name references to '$( $newimageName ): $( $_.Exception.Message )"
                }
            }

            # Mutate markdown content
            try {
                if ($config['dryRun']['value'] -eq 1) {
                    # Get markdown content
                    $content = @( Get-Content -Path "$( $pageCfg['fullfilepathwithoutextension'] ).md" -ErrorAction Stop )  # Get-Content -ErrorAction Stop can produce random "Cannot find path 'xxx' because it does not exist"
                    $content = @(
                        if ($content.Count -gt 6) {
                            # Discard first 6 lines which contain a header, created date, and time. We are going to add our own header
                            $content[6..($content.Count - 1)]
                        }else {
                            # Empty page
                            ''
                        }
                    ) -join "`r`n"
                }

                # Mutate
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        try {
                            "Mutation of markdown: $( $m['description'] )" | Write-Verbose
                            if ($config['dryRun']['value'] -eq 1) {
                                $content = $content -replace $r['searchRegex'], $r['replacement']
                            }
                        }catch {
                            Write-Error "Failed to mutating markdown content with mutation '$( $m['description'] )': $( $_.Exception.Message )"
                        }
                    }
                }
                if ($config['dryRun']['value'] -eq 1) {
                    Set-Content "$( $pageCfg['fullfilepathwithoutextension'] ).md" -Value $content -ErrorAction Stop
                }
            }catch {
                Write-Error "Error while mutating markdown content: $( $_.Exception.Message )"
            }

            "Markdown file ready: $( $pageCfg['fullfilepathwithoutextension'] ).md" | Write-Host -ForegroundColor Green
        }catch {
            Write-Error "Failed to convert page: $( $pageCfg['pathFromRoot'] ). Reason: $( $_.Exception.Message )"
        }
    }
}

# Unused
Function Convert-OneNoteSectionGroup {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $OneNoteConnection
    ,
        # ConvertOneNote2MarkDown configuration object
        [Parameter(Mandatory)]
        [object]
        $Config
    ,
        # Conversion object
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object]
        $ConversionConfig
    ,
        [Parameter()]
        [switch]
        $Recurse

    )

    foreach($cfg in $ConversionConfig) {
        if ($cfg['levelsFromRoot'] -eq 0) {
            "Notebook: $( $cfg.object.name )" | Write-Host -ForegroundColor Green
        }
        foreach ($sectionCfg in $cfg['sections']) {
            "$( '#' * ($sectionCfg['levelsFromRoot']) ) $( $sectionCfg['object'].name ) [$( $sectionCfg['kind'] )]" | Write-Host
            "Uri: $( $sectionCfg['uri'] ) [$( $sectionCfg['kind'] )]" | Write-Verbose
            foreach ($pageCfg in $sectionCfg['pages']) {
                Convert-OneNotePage -OneNoteConnection $OneNoteConnection -Config $Config -ConversionConfig $pageCfg
            }
        }

        if ($Recurse) {
            foreach ($sectionGroupCfg in $cfg['sectionGroups']) {
                "`n$( '#' * ($sectionGroupCfg['levelsFromRoot']) ) $( $sectionGroupCfg['object'].name ) [$( $sectionGroupCfg['kind'] )]" | Write-Host
                "Uri: $( $sectionGroupCfg['uri'] )" | Write-Verbose
                Convert-OneNoteSectionGroup -OneNoteConnection $OneNoteConnection -Config $Config -ConversionConfig $sectionGroupCfg -Recurse
            }
        }
    }
}

Function Print-ConversionErrors {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorCollection
    )

    if ($ErrorCollection.Count -gt 0) {
        "Conversion errors: " | Write-Host
        $ErrorCollection | Where-Object { (Get-Member -InputObject $_ -Name 'CategoryInfo') -and ($_.CategoryInfo.Reason -match 'WriteErrorException') } | Write-Host
    }
}

Function Convert-OneNote2MarkDown {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $ConversionConfigurationExportPath
    )

    try {
        Set-StrictMode -Version Latest

        # Fix encoding problems for languages other than English
        $PSDefaultParameterValues['*:Encoding'] = 'utf8'

        $totalerr = @()

        # Validate dependencies
        Validate-Dependencies

        # Compile and validate configuration
        $config = Compile-Configuration | Validate-Configuration

        "Configuration:" | Write-Host -ForegroundColor Cyan
        $config | Print-Configuration

        # Connect to OneNote
        $OneNote = New-OneNoteConnection

        # Get the hierarchy of OneNote objects as xml
        $hierarchy = Get-OneNoteHierarchy -OneNoteConnection $OneNote

        # Get and validate the notebook(s) to convert
        $notebooks = @(
            if ($config['targetNotebook']['value']) {
                $hierarchy.Notebooks.Notebook | Where-Object { $_.Name -match $config['targetNotebook']['value'] }
            }else {
                $hierarchy.Notebooks.Notebook
            }
        )
        if ($notebooks.Count -eq 0) {
            if ($config['targetNotebook']['value']) {
                throw "Could not find notebook of name '$( $config['targetNotebook']['value'] )'"
            }else {
                throw "Could not find notebooks"
            }
        }

        # Convert the notebook(s)
        "`nConverting notes..." | Write-Host -ForegroundColor Cyan
        New-SectionGroupConversionConfig -OneNoteConnection $OneNote -NotesDestination $config['notesdestpath']['value'] -Config $config -SectionGroups $notebooks -LevelsFromRoot 0 -ErrorVariable +totalerr | Tee-Object -Variable pageConversionConfigs | Convert-OneNotePage -OneNoteConnection $OneNote -Config $config -ErrorVariable +totalerr
        "Done converting notes." | Write-Host -ForegroundColor Cyan

        # Export all Page Conversion Configuration objects as .json, which is useful for debugging
        if ($ConversionConfigurationExportPath) {
            "Exporting Page Conversion Configuration as JSON file: $ConversionConfigurationExportPath" | Write-Host -ForegroundColor Cyan
            $pageConversionConfigs | ConvertTo-Json -Depth 100 | Out-File $ConversionConfigurationExportPath -Encoding utf8 -Force
        }
    }catch {
        if ($ErrorActionPreference -eq 'Stop') {
            throw
        }else {
            Write-Error -ErrorRecord $_
        }
    }finally {
        'Cleaning up...' | Write-Host -ForegroundColor Cyan

        # Disconnect OneNote connection
        if (Get-Variable -Name OneNote -ErrorAction SilentlyContinue) {
            Remove-OneNoteConnection -OneNoteConnection $OneNote
        }

        # Print any conversion errors
        Print-ConversionErrors -ErrorCollection $totalerr
        'Exiting.' | Write-Host -ForegroundColor Cyan
    }
}

if (!$Exit) {
    # Entrypoint
    $params = @{
        ConversionConfigurationExportPath = $ConversionConfigurationExportPath
    }
    Convert-OneNote2MarkDown @params
}
