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

    # Validate Powershell versions. Supported version are between 5.x (possibly lower) and 7.0.x
    if ($PSVersionTable.PSVersion -ge [version]'7.1') {
        throw "Unsupported Powershell version $( $PSVersionTable.PSVersion ), because it does not support importing Win32 GAC Assemblies. Supported versions are between Powershell 5.x and 7.0.x. See README.md for instructions to install Powershell 7.0.x"
    }

    # Validate assemblies
    if ( ($env:OS -imatch 'Windows') -and ! (Get-Item -Path $env:windir\assembly\GAC_MSIL\*onenote*) ) {
        "There are missing onenote assemblies. Please ensure the Desktop version of Onenote 2016 or above is installed." | Write-Warning
    }

    # Validate dependencies
    if (! (Get-Command -Name 'pandoc.exe' -ErrorAction SilentlyContinue) ) {
        throw "Could not locate pandoc.exe. Please ensure pandoc is installed for all users, and available in PATH. If pandoc was just installed using .msi or chocolatey, you may need to restart Powershell or the computer for pandoc to be set correctly in PATH."
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
0: Convert - Default
1: Dry run
'@
            default = 0
            value = 0
            validateRange = 0,1
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
        docxNamingConvention = @{
            description = @'
Whether to use name .docx files using page ID with last modified date epoch, or hierarchy
1: Use page ID with last modified date epoch (recommended if you chose to use existing .docx files) - Default
2: Use hierarchy
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
        mdFileNameAndFolderNameMaxLength = @{
            description = @'
Specify a value between 32 and 255 as the maximum length of markdown file names, and their folder names (only when using subfolders for subpages (e.g. Page\Subpage.md)). File and folder names with length exceeding this value will be truncated accordingly.
NOTE: If you are using prefixes for subpages (e.g. Page_Subpage.md), it is recommended to set this to at 100 or more.
Default: 32
'@
            default = 32
            value = 32
            validateRange = 32,255
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
        mergeInkStrokes = @{
            description = @'
Whether to merge the strokes of pen/hand written drawings (InkDrawing) into a single image that preserves the position of each stroke. Strokes drawn on an image are drawn on that image. Requires Windows
1: Don't merge. Each stroke is converted to a separate image - Default
2: Merge strokes into a single image
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        conversion = @{
            description = @'
Specify Pandoc output format and optional extension(s) in the format: <markdownformat><+extension><-extension>
Use this to customize the markdown flavor
Extension(s) must be supported by your pandoc version. See: https://pandoc.org/MANUAL.html#options
Examples:
  markdown-simple_tables-multiline_tables-grid_tables+pipe_tables
  commonmark+pipe_tables
  gfm+pipe_tables
  markdown_mmd-simple_tables-multiline_tables-grid_tables+pipe_tables
  markdown_phpextra-simple_tables-multiline_tables-grid_tables+pipe_tables
  markdown_strict+simple_tables-multiline_tables-grid_tables+pipe_tables
Default:
  markdown-simple_tables-multiline_tables-grid_tables+pipe_tables
'@
            default = 'markdown-simple_tables-multiline_tables-grid_tables+pipe_tables'
            value = 'markdown-simple_tables-multiline_tables-grid_tables+pipe_tables'
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
Whether to clear extra newlines between unordered (bullet) and ordered (numbered) list items, non-breaking spaces from blank lines, and `>` after unordered lists
1: Clear - Default
2: Don't clear
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        keepescape = @{
            description = @'
Whether to clear escape symbols from md files. See: https://pandoc.org/MANUAL.html#backslash-escapes
1: Clear all '\' characters  - Default
2: Clear all '\' characters except those preceding alphanumeric characters
3: Keep '\' symbol escape
'@
            default = 1
            value = 1
            validateRange = 1,3
        }
        newlineCharacter = @{
            description = @'
Whether to use Line Feed (LF) or Carriage Return + Line Feed (CRLF) for new lines
1: LF (unix) - Default
2: CRLF (windows)
'@
            default = 1
            value = 1
            validateRange = 1,2
        }
        exportPdf = @{
            description = @'
Whether to include a PDF export alongside the markdown file
1: Don't include PDF - Default
2: Include PDF
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
    $configFile = [io.path]::combine( $PSScriptRoot, 'config.ps1' )
    if (Test-Path $configFile) {
        try {
            & {
                $scriptblock = [scriptblock]::Create( (Get-Content -LiteralPath $configFile -Raw) )
                . $scriptblock *>$null # Cleanup the pipeline
                foreach ($key in @($config.Keys)) {
                    # E.g. 'string', 'int'
                    $typeName = [Microsoft.PowerShell.ToStringCodeMethods]::Type($config[$key]['default'].GetType())
                    $config[$key]['value'] = Invoke-Expression -Command "(Get-Variable -Name `$key -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Value) -as [$typeName]"
                    if ($config[$key]['value'] -is [string]) {
                        # Trim string
                        $config[$key]['value'] = $config[$key]['value'].Trim()

                        # Remove trailing slash(es) for paths
                        if ($key -match 'path' -and $config[$key]['value'] -match '[/\\]') {
                            $config[$key]['value'] = $config[$key]['value'].TrimEnd('/').TrimEnd('\')
                        }
                    }
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
        }catch {
            Write-Warning "There is an error in the configuration file $configFile $( $_.ScriptStackTrace ). `nThe exception was: $( $_.Exception.Message )"
            throw
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
            if (! $Config.Contains($key) -or ($null -eq $Config[$key]) -or ($null -eq $Config[$key]['value'])) {
                throw "Missing or invalid configuration option '$key'. Expected a value of type $( $defaultConfig[$key]['default'].GetType().FullName )"
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

Function Truncate-PathFileName {
    [CmdletBinding(DefaultParameterSetName='default')]
    param (
        [Parameter(ParameterSetName='default',Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path
    ,
        [Parameter(ParameterSetName='pipeline',ValueFromPipeline)]
        [string]
        $InputObject
    ,
        [Parameter()]
        [ValidateRange(0,255)]
        [int]
        $Length
    )

    process {
        if ($InputObject) {
            $Path = $InputObject
        }
        if ($null -eq $Path) {
            throw "No input parameters specified."
        }
        $maxLength = 255
        if ($Length) {
            $maxLength = $Length
        }

        # On Windows, even with support for long absolute file paths, there's still a limit for file or folder names (i.e. File or folder name limit: Max 255 characters long)
        $name = Split-Path $Path -Leaf
        if ($name.Length -gt $maxLength) {
            $parent = Split-Path $Path -Parent
            $truncatedName = $name.Substring(0, $maxLength)
            [io.path]::combine( $parent, $truncatedName )
        }else {
            $Path
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

Function Encode-Markdown {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]
        $Name
    ,
        [Parameter()]
        [switch]
        $Uri
    )

    if ($Uri) {
        $markdownChars = '[]()'.ToCharArray()
        foreach ($c in $markdownChars) {
            $Name = $Name.Replace("$c", "\$c")
        }
    }else {
        # See: https://pandoc.org/MANUAL.html#backslash-escapes
        $markdownChars = '\*_{}[]()#+-.!'.ToCharArray()
        foreach ($c in $markdownChars) {
            $Name = $Name.Replace("$c", "\$c")
        }
        $markdownChars2 = '`'
        foreach ($c in $markdownChars2) {
            $Name = $Name.Replace("$c", "$c$c$c")
        }
    }
    $Name
}

Function Set-ContentNoBom {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline = $true,ValueFromPipelineByPropertyName = $true)]
        [AllowEmptyString()]
        [string]
        $LiteralPath
    ,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [array]
        $Value
    )
    process {
        if ($PSVersionTable.PSVersion.Major -le 5) {
            try {
                $content = $Value -join ''
                [IO.File]::WriteAllLines($LiteralPath, $content)
            }catch {
                if ($ErrorActionPreference -eq 'Stop') {
                    throw
                }else {
                    Write-Error -ErrorRecord $_
                }
            }
        }else {
            Set-Content @PSBoundParameters
        }
    }
}

Function New-OneNoteConnection {
    [CmdletBinding()]
    param ()

    # Create a OneNote connection. See: https://docs.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote
    if ($PSVersionTable.PSVersion.Major -le 5) {
        if ($OneNote = New-Object -ComObject OneNote.Application) {
            $OneNote
        }else {
            Write-Error "Failed to make connection to OneNote." -ErrorAction Continue
            throw
        }
    }else {
        # Works between powershell 5.x (possibly lower) and 7.0, but not >= 7.1. 7.1 and above doesn't seem to support loading Win32 GAC Assemblies.
        if (Add-Type -Path $env:windir\assembly\GAC_MSIL\Microsoft.Office.Interop.OneNote\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.OneNote.dll -PassThru) {
            $OneNote = [Microsoft.Office.Interop.OneNote.ApplicationClass]::new()
            $OneNote
        }else {
            Write-Error "Failed to make connection to OneNote." -ErrorAction Continue
            throw
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

Function Publish-OneNotePage {
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
    ,
        [Parameter(Mandatory)]
        [ValidateSet('pfOneNotePackage', 'pfOneNotePackage', 'pfOneNote ', 'pfPDF', 'pfXPS', 'pfWord', 'pfEMF', 'pfHTML', 'pfOneNote2007')]
        [ValidateNotNullOrEmpty()]
        [string]
        $PublishFormat
    )

    $OneNoteConnection.Publish($PageId, $Destination, $PublishFormat, "")
}

Function Get-DocxImages {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Path
    )

    # Get the images of a .docx file in document order, with their file names and display sizes in points
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entries = @{}
        foreach ($name in 'word/document.xml', 'word/_rels/document.xml.rels') {
            $reader = New-Object System.IO.StreamReader($zip.GetEntry($name).Open())
            try {
                $entries[$name] = [xml]$reader.ReadToEnd()
            }finally {
                $reader.Dispose()
            }
        }
    }finally {
        $zip.Dispose()
    }

    $relationshipsNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $targets = @{}
    foreach ($r in $entries['word/_rels/document.xml.rels'].DocumentElement.SelectNodes('*')) {
        $targets[$r.GetAttribute('Id')] = $r.GetAttribute('Target')
    }
    $document = $entries['word/document.xml']
    $ns = New-Object Xml.XmlNamespaceManager $document.NameTable
    $ns.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
    $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    foreach ($drawing in $document.SelectNodes('//wp:inline | //wp:anchor', $ns)) {
        $extent = $drawing.SelectSingleNode('wp:extent', $ns)
        $blip = $drawing.SelectSingleNode('.//a:blip', $ns)
        if ($extent -and $blip -and $targets[$blip.GetAttribute('embed', $relationshipsNs)]) {
            [ordered]@{
                fileName = Split-Path $targets[$blip.GetAttribute('embed', $relationshipsNs)] -Leaf
                width = [double]$extent.GetAttribute('cx') / 12700 # EMUs to points
                height = [double]$extent.GetAttribute('cy') / 12700 # EMUs to points
            }
        }
    }
}

Function Get-OneNotePageImages {
    [CmdletBinding()]
    param (
        # Page content
        [Parameter(Mandatory)]
        [xml]
        $PageXml
    )

    # Get the images of a page in the same order as in the page's .docx export (i.e. outlines and images from top to bottom), with their sizes in points.
    # Also get their positions on the page in points, where they can be determined. Images in outlines have no position, so it is determined from the outline's position and the heights of the paragraphs preceding or following the image, if they are known (i.e. images and empty paragraphs). Otherwise, the position is $null
    $ns = New-Object Xml.XmlNamespaceManager $PageXml.NameTable
    $ns.AddNamespace("one", $PageXml.DocumentElement.NamespaceURI)
    $quickStyles = @{}
    foreach ($q in $PageXml.SelectNodes('/one:Page/one:QuickStyleDef', $ns)) {
        $quickStyles[$q.index] = $q
    }
    $lineSpacings = @{}
    $objects = @( $PageXml.SelectNodes('/one:Page/one:Outline[one:Position] | /one:Page/one:Image[one:Position and one:Size]', $ns) | Sort-Object { [double]$_.Position.y } )
    foreach ($o in $objects) {
        if ($o.LocalName -eq 'Image') {
            [ordered]@{
                x = [double]$o.Position.x
                y = [double]$o.Position.y
                width = [double]$o.Size.width
                height = [double]$o.Size.height
            }
            continue
        }

        # Determine the heights of the outline's paragraphs. -1 if unknown
        $paragraphs = @( $o.SelectNodes('one:OEChildren/one:OE', $ns) )
        $heights = New-Object double[] $paragraphs.Count
        for ($i = 0; $i -lt $paragraphs.Count; $i++) {
            $p = $paragraphs[$i]
            $children = @( $p.SelectNodes('*') )
            $heights[$i] = -1
            if ($children.Count -ne 1 -or $p.HasAttribute('style') -or $p.HasAttribute('spaceBefore') -or $p.HasAttribute('spaceAfter') -or $p.HasAttribute('spaceBetween')) {
                continue
            }
            if ($children[0].LocalName -eq 'Image' -and $children[0].SelectSingleNode('one:Size', $ns)) {
                $heights[$i] = [double]$children[0].Size.height
            }elseif ($children[0].LocalName -eq 'T' -and $children[0].InnerText.Trim() -eq '' -and $quickStyles.ContainsKey($p.GetAttribute('quickStyleIndex'))) {
                # Empty paragraph, i.e. a single line in the paragraph's font
                $q = $quickStyles[$p.GetAttribute('quickStyleIndex')]
                if (!$lineSpacings.ContainsKey($q.font)) {
                    $lineSpacings[$q.font] = try {
                        Add-Type -AssemblyName PresentationCore, WindowsBase -ErrorAction Stop # Not available on non-Windows
                        $glyphTypeface = $null
                        if ((New-Object System.Windows.Media.Typeface($q.font) -ErrorAction Stop).TryGetGlyphTypeface([ref]$glyphTypeface)) {
                            (New-Object System.Windows.Media.FontFamily($q.font) -ErrorAction Stop).LineSpacing
                        }else {
                            -1 # Font is not installed
                        }
                    }catch {
                        -1
                    }
                }
                if ($lineSpacings[$q.font] -gt 0) {
                    $heights[$i] = [double]$q.fontSize * $lineSpacings[$q.font] + [double]$q.spaceBefore + [double]$q.spaceAfter
                }
            }
        }

        # Determine the top of the paragraphs, from the top of the outline downwards, and from the bottom of the outline upwards
        $tops = @{}
        $top = [double]$o.Position.y
        for ($i = 0; $i -lt $paragraphs.Count; $i++) {
            $tops[$i] = $top
            if ($heights[$i] -lt 0) {
                break
            }
            $top += $heights[$i]
        }
        $size = $o.SelectSingleNode('one:Size', $ns)
        if ($size) {
            $bottom = [double]$o.Position.y + [double]$size.height
            for ($i = $paragraphs.Count - 1; $i -ge 0; $i--) {
                if ($heights[$i] -lt 0) {
                    break
                }
                $bottom -= $heights[$i]
                $tops[$i] = $bottom
            }
            # Validate: The known heights cannot exceed the outline's height, and must add up to the outline's height if all of them are known
            $knownHeights = ( $heights | Where-Object { $_ -ge 0 } | Measure-Object -Sum ).Sum
            if ($knownHeights -gt [double]$size.height + 1 -or (@( $heights | Where-Object { $_ -lt 0 } ).Count -eq 0 -and [Math]::Abs($knownHeights - [double]$size.height) -gt 1)) {
                "Could not determine the layout of an outline. Expected height: $( $size.height ), calculated height: $knownHeights" | Write-Verbose
                $tops = @{}
            }
        }

        # Get the images in document order. Only images in the outline's top level paragraphs have a determinable position
        foreach ($image in $o.SelectNodes('.//one:Image[one:Size]', $ns)) {
            $imageCfg = [ordered]@{
                x = $null
                y = $null
                width = [double]$image.Size.width
                height = [double]$image.Size.height
            }
            $i = [array]::IndexOf($paragraphs, $image.ParentNode)
            if ($i -ge 0 -and $tops.ContainsKey($i) -and $size) {
                $imageCfg['x'] = switch ($image.ParentNode.GetAttribute('alignment')) {
                    'center' { [double]$o.Position.x + ([double]$size.width - $imageCfg['width']) / 2 }
                    'right' { [double]$o.Position.x + [double]$size.width - $imageCfg['width'] }
                    default { [double]$o.Position.x }
                }
                $imageCfg['y'] = $tops[$i]
            }
            $imageCfg
        }
    }
}

Function New-InkDrawingImage {
    [CmdletBinding()]
    param (
        # Ink drawing configuration objects
        [Parameter(Mandatory)]
        [object[]]
        $InkDrawing
    ,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Destination
    ,
        # A .png or .jpg image that was drawn on. The strokes are rendered on top of it
        [Parameter()]
        [string]
        $BackgroundImagePath
    ,
        # Position and size of the image that was drawn on, in points
        [Parameter()]
        [object]
        $BackgroundImageConfig
    )

    # Render the strokes of the ink drawing(s) into a single image, preserving the position of each stroke. Returns the size of the image in points
    Add-Type -AssemblyName PresentationCore, WindowsBase
    $strokes = New-Object System.Windows.Ink.StrokeCollection
    foreach ($i in $InkDrawing) {
        $stream = New-Object System.IO.MemoryStream(,[Convert]::FromBase64String($i['data']))
        $strokes.Add( (New-Object System.Windows.Ink.StrokeCollection -ArgumentList (,$stream)) ) # Stroke coordinates are relative to the page, in device-independent pixels (1/96 inch)
    }
    if ($strokes.Count -eq 0) {
        throw "No strokes found in ink drawing(s)"
    }
    $bounds = $strokes.GetBounds()
    if ($BackgroundImagePath) {
        # Keep the resolution of the image. Extend the image if any stroke exceeds it
        $background = [System.Windows.Media.Imaging.BitmapFrame]::Create( (New-Object System.IO.MemoryStream(,[System.IO.File]::ReadAllBytes($BackgroundImagePath))), [System.Windows.Media.Imaging.BitmapCreateOptions]::None, [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad )
        $backgroundRect = New-Object System.Windows.Rect(($BackgroundImageConfig['x'] / 0.75), ($BackgroundImageConfig['y'] / 0.75), ($BackgroundImageConfig['width'] / 0.75), ($BackgroundImageConfig['height'] / 0.75)) # Points to pixels
        $scaleX = $background.PixelWidth / $backgroundRect.Width
        $scaleY = $background.PixelHeight / $backgroundRect.Height
        $bounds = [System.Windows.Rect]::Union($bounds, $backgroundRect)
        # Align the extension to whole pixels of the image, so that the image is not resampled
        $left = $backgroundRect.Left - [Math]::Ceiling(($backgroundRect.Left - $bounds.Left) * $scaleX) / $scaleX
        $top = $backgroundRect.Top - [Math]::Ceiling(($backgroundRect.Top - $bounds.Top) * $scaleY) / $scaleY
        $right = $backgroundRect.Right + [Math]::Ceiling(($bounds.Right - $backgroundRect.Right) * $scaleX) / $scaleX
        $bottom = $backgroundRect.Bottom + [Math]::Ceiling(($bounds.Bottom - $backgroundRect.Bottom) * $scaleY) / $scaleY
        $bounds = New-Object System.Windows.Rect($left, $top, ($right - $left), ($bottom - $top))
    }else {
        $padding = 2
        $bounds = New-Object System.Windows.Rect(($bounds.X - $padding), ($bounds.Y - $padding), ($bounds.Width + 2 * $padding), ($bounds.Height + 2 * $padding))
        $scaleX = $scaleY = 2 # Render at 192 DPI so that the image stays sharp
    }
    $visual = New-Object System.Windows.Media.DrawingVisual
    $context = $visual.RenderOpen()
    $context.PushTransform( (New-Object System.Windows.Media.MatrixTransform($scaleX, 0, 0, $scaleY, (-$bounds.X * $scaleX), (-$bounds.Y * $scaleY))) )
    $context.DrawRectangle([System.Windows.Media.Brushes]::White, $null, $bounds)
    if ($BackgroundImagePath) {
        $context.DrawImage($background, $backgroundRect)
    }
    $strokes.Draw($context)
    $context.Pop()
    $context.Close()
    $bitmap = New-Object System.Windows.Media.Imaging.RenderTargetBitmap([int][Math]::Round($bounds.Width * $scaleX), [int][Math]::Round($bounds.Height * $scaleY), 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($visual)
    $encoder = if ($Destination -match '\.jpe?g$') {
        New-Object System.Windows.Media.Imaging.JpegBitmapEncoder -Property @{ QualityLevel = 95 }
    }else {
        New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    }
    $encoder.Frames.Add( [System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap) )
    $fileStream = [System.IO.File]::Create($Destination)
    try {
        $encoder.Save($fileStream)
    }finally {
        $fileStream.Dispose()
    }

    @{
        width = $bounds.Width * 0.75 # Pixels to points
        height = $bounds.Height * 0.75 # Pixels to points
    }
}

Function Merge-InkDrawingStrokes {
    [CmdletBinding()]
    param (
        # Page conversion configuration object
        [Parameter(Mandatory)]
        [object]
        $PageConfig
    )

    $pageCfg = $PageConfig
    $docxImages = @( Get-DocxImages -Path $pageCfg['docxExportFilePath'] )

    # Identify the stroke images of the .docx. Each ink drawing is exported as a separate image, in the same order as in the page, sized as the ink drawing plus a padding of about 8.25 points
    $inkDrawingImages = @{}
    $i = 0
    foreach ($image in $docxImages) {
        if ($i -lt $pageCfg['inkDrawings'].Count) {
            $inkDrawingCfg = $pageCfg['inkDrawings'][$i]
            if ([Math]::Abs($image['width'] - $inkDrawingCfg['width'] - 8.25) -le 1.5 -and [Math]::Abs($image['height'] - $inkDrawingCfg['height'] - 8.25) -le 1.5) {
                $inkDrawingImages[$image['fileName']] = $inkDrawingCfg
                $i++
            }
        }
    }
    "Identified $( $inkDrawingImages.Count ) of $( $pageCfg['inkDrawings'].Count ) ink drawing stroke images" | Write-Verbose
    if ($inkDrawingImages.Count -eq 0) {
        return
    }

    # Identify the other images of the .docx. They are exported in the same order as in the page, with the same size
    $images = @{}
    $pageImages = @( $pageCfg['images'] | Where-Object { $null -ne $_ } )
    $i = 0
    foreach ($image in $docxImages) {
        if ($i -lt $pageImages.Count -and !$inkDrawingImages.ContainsKey($image['fileName'])) {
            $imageCfg = $pageImages[$i]
            if ([Math]::Abs($image['width'] - $imageCfg['width']) -le 1.5 -and [Math]::Abs($image['height'] - $imageCfg['height']) -le 1.5) {
                $images[$image['fileName']] = $imageCfg
                $i++
            }
        }
    }

    $content = Get-Content -LiteralPath $pageCfg['filePath'] -Raw -ErrorAction Stop # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
    if (!$content) {
        return
    }
    $imageRegex = "!\[(?:\\.|[^\\\]])*\]\($( [regex]::Escape($pageCfg['mediaPathPandoc']) )/([^)\s]+)\)(\{[^}\r\n]*\})?"
    $references = @( [regex]::Matches($content, $imageRegex) )
    $referencedFileNames = @{}
    foreach ($r in $references) {
        $referencedFileNames[$r.Groups[1].Value] = $true
    }

    # Assign each stroke drawn on an image to that image, i.e. if the stroke's center is within an image with a known position. Only .png and .jpg images that are referenced in the markdown are supported
    $drawnOnImages = @{}
    $strokesOnImages = @{}
    foreach ($image in $docxImages) {
        if (!$inkDrawingImages.ContainsKey($image['fileName'])) {
            continue
        }
        $inkDrawingCfg = $inkDrawingImages[$image['fileName']]
        $centerX = $inkDrawingCfg['x'] + $inkDrawingCfg['width'] / 2
        $centerY = $inkDrawingCfg['y'] + $inkDrawingCfg['height'] / 2
        foreach ($imageFileName in $images.Keys) {
            $imageCfg = $images[$imageFileName]
            if ($null -ne $imageCfg['x'] -and $null -ne $imageCfg['y'] -and $imageFileName -match '\.(png|jpe?g)$' -and $referencedFileNames.ContainsKey($imageFileName) -and
                $centerX -ge $imageCfg['x'] -and $centerX -le $imageCfg['x'] + $imageCfg['width'] -and $centerY -ge $imageCfg['y'] -and $centerY -le $imageCfg['y'] + $imageCfg['height']) {
                if (!$drawnOnImages.ContainsKey($imageFileName)) {
                    $drawnOnImages[$imageFileName] = [System.Collections.ArrayList]@()
                }
                $drawnOnImages[$imageFileName].Add($image['fileName']) > $null
                $strokesOnImages[$image['fileName']] = $true
                break
            }
        }
    }

    # Render the strokes on the images they were drawn on. If that fails, the strokes are merged into a separate image instead
    $drawnOnImageSizes = @{}
    foreach ($imageFileName in $drawnOnImages.Keys) {
        $imagePath = [io.path]::combine( $pageCfg['mediaPathPandoc'], $imageFileName )
        try {
            "Drawing $( $drawnOnImages[$imageFileName].Count ) strokes on image $imagePath" | Write-Verbose
            $drawnOnImageSizes[$imageFileName] = New-InkDrawingImage -InkDrawing @( $drawnOnImages[$imageFileName] | ForEach-Object { $inkDrawingImages[$_] } ) -Destination $imagePath -BackgroundImagePath $imagePath -BackgroundImageConfig $images[$imageFileName]
        }catch {
            Write-Error "Failed to draw strokes on image $imagePath. Exception: $( $_.Exception.Message )" -ErrorAction Continue
            foreach ($f in $drawnOnImages[$imageFileName]) {
                $strokesOnImages.Remove($f)
            }
        }
    }

    # Group adjacent stroke image references, e.g. '![](<mediaPathPandoc>/image2.png)![](<mediaPathPandoc>/image3.png)'
    $groups = [System.Collections.ArrayList]@()
    $edits = [System.Collections.ArrayList]@()
    $group = $null
    foreach ($r in $references) {
        $fileName = $r.Groups[1].Value
        if ($inkDrawingImages.ContainsKey($fileName)) {
            if (!$group -or $content.Substring($group['end'], $r.Index - $group['end']) -notmatch '^\s*$') {
                $group = @{
                    start = $r.Index
                    fileNames = [System.Collections.ArrayList]@()
                    hasAttributes = $false
                }
                $groups.Add($group) > $null
            }
            $group['end'] = $r.Index + $r.Length
            if (!$strokesOnImages.ContainsKey($fileName)) {
                $group['fileNames'].Add($fileName) > $null
                $group['hasAttributes'] = $group['hasAttributes'] -or $r.Groups[2].Success
            }
        }else {
            $group = $null
            if ($drawnOnImageSizes.ContainsKey($fileName) -and $r.Groups[2].Success) {
                # Update the size of an image that was drawn on, if strokes exceed the image
                $size = $drawnOnImageSizes[$fileName]
                if ([Math]::Abs($size['width'] - $images[$fileName]['width']) -gt 0.5 -or [Math]::Abs($size['height'] - $images[$fileName]['height']) -gt 0.5) {
                    $edits.Add(@{
                        start = $r.Groups[2].Index
                        end = $r.Groups[2].Index + $r.Groups[2].Length
                        replacement = "{width=`"$( [Math]::Round($size['width'] / 72, 4) )in`" height=`"$( [Math]::Round($size['height'] / 72, 4) )in`"}" # Points to inches
                    }) > $null
                }
            }
        }
    }

    # Replace each group with a single image of the strokes that were not drawn on an image, or remove the group if there are none
    $fileNamesToRemove = @( $strokesOnImages.Keys )
    foreach ($group in $groups) {
        $replacement = ''
        if ($group['fileNames'].Count -gt 0) {
            $fileName = $group['fileNames'][0]
            $size = New-InkDrawingImage -InkDrawing @( $group['fileNames'] | ForEach-Object { $inkDrawingImages[$_] } ) -Destination ([io.path]::combine( $pageCfg['mediaPathPandoc'], $fileName ))
            $fileNamesToRemove += @( $group['fileNames'] | Select-Object -Skip 1 )
            $replacement = "![]($( $pageCfg['mediaPathPandoc'] )/$fileName)"
            if ($group['hasAttributes']) {
                $replacement += "{width=`"$( [Math]::Round($size['width'] / 72, 4) )in`" height=`"$( [Math]::Round($size['height'] / 72, 4) )in`"}" # Points to inches
            }
        }
        $edits.Add(@{
            start = $group['start']
            end = $group['end']
            replacement = $replacement
        }) > $null
    }

    # Apply the edits. Begin from the last edit, so that the positions of preceding edits stay intact
    foreach ($e in @( $edits | Sort-Object { $_.start } -Descending )) { # Use dot notation, since index notation fails in Sort-Object in Powershell 6 and 7
        $content = $content.Substring(0, $e['start']) + $e['replacement'] + $content.Substring($e['end'])
    }
    foreach ($f in $fileNamesToRemove) {
        Remove-Item -LiteralPath ([io.path]::combine( $pageCfg['mediaPathPandoc'], $f )) -Force -ErrorAction SilentlyContinue
    }
    Set-ContentNoBom -LiteralPath $pageCfg['filePath'] -Value $content -ErrorAction Stop # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
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
        $cfg['id'] = $sectionGroup.ID # E.g. {9570CCF6-17C2-4DCE-83A0-F58AE8914E29}{1}{B0}
        $cfg['nameCompat'] = $sectionGroup.name | Remove-InvalidFileNameChars
        $cfg['levelsFromRoot'] = $LevelsFromRoot
        $cfg['uri'] = $sectionGroup.path # E.g. https://d.docs.live.net/0123456789abcdef/Skydrive Notebooks/mynotebook/mysectiongroup
        $cfg['notesDirectory'] = [io.path]::combine( $NotesDestination.TrimEnd('/').TrimEnd('\').Replace('\', [io.path]::DirectorySeparatorChar), $cfg['nameCompat'] ) # No need to truncate. Section Group and Section names have a max length of 50, so we should never hit the absolute path, file name, or directory name limits on Windows
        $cfg['notesBaseDirectory'] = & {
            # E.g. 'c:\temp\notes\mynotebook\mysectiongroup'
            # E.g. levelsFromRoot: 1
            $split = $cfg['notesDirectory'].Split( [io.path]::DirectorySeparatorChar )
            # E.g. 5
            $totalLevels = $split.Count
            # E.g. 0..(5-1-1) -> 'c:\temp\notes\mynotebook'
            $split[0..($totalLevels - $cfg['levelsFromRoot'] - 1)] -join [io.path]::DirectorySeparatorChar
        }
        $cfg['notebookName'] = Split-Path $cfg['notesBaseDirectory'] -Leaf
        $cfg['pathFromRoot'] = $cfg['notesDirectory'].Replace($cfg['notesBaseDirectory'], '').Trim([io.path]::DirectorySeparatorChar)
        $cfg['pathFromRootCompat'] = $cfg['pathFromRoot'] | Remove-InvalidFileNameChars
        $cfg['notesDocxDirectory'] = [io.path]::combine( $cfg['notesBaseDirectory'], 'docx' )
        $cfg['directoriesToCreate'] = @()

        # Build this Section Group's sections
        $cfg['sections'] = [System.Collections.ArrayList]@()

        if (! (Get-Member -InputObject $sectionGroup -Name 'Section') -and ! (Get-Member -InputObject $sectionGroup -Name 'SectionGroup') ) {
            "Ignoring empty Section Group: $( $cfg['pathFromRoot'] )" | Write-Host -ForegroundColor DarkGray
        }

        if (Get-Member -InputObject $sectionGroup -Name 'Section') {
            foreach ($section in $sectionGroup.Section) {
                "$( '#' * ($LevelsFromRoot + 1) ) Building conversion configuration for $( $section.name ) [Section]" | Write-Host -ForegroundColor DarkGray

                $sectionCfg = [ordered]@{}
                $sectionCfg['notebookName'] = $cfg['notebookName']
                $sectionCfg['notesBaseDirectory'] = $cfg['notesBaseDirectory']
                $sectionCfg['notesDirectory'] = $cfg['notesDirectory']
                $sectionCfg['sectionGroupUri'] = $cfg['uri'] # Keep a reference to my Section Group Configuration object's uri
                $sectionCfg['sectionGroupName'] = $cfg['object'].name
                $sectionCfg['object'] = $section # Keep a reference to the Section object
                $sectionCfg['kind'] = 'Section'
                $sectionCfg['id'] = $section.ID # E.g {BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{B0}
                $sectionCfg['nameCompat'] = $section.name | Remove-InvalidFileNameChars
                $sectionCfg['levelsFromRoot'] = $cfg['levelsFromRoot'] + 1
                $sectionCfg['pathFromRoot'] = "$( $cfg['pathFromRoot'] )$( [io.path]::DirectorySeparatorChar )$( $sectionCfg['nameCompat'] )".Trim([io.path]::DirectorySeparatorChar) # No need to truncate. Section Group and Section names have a max length of 50, so we should never hit the absolute path, file name, or directory name limits on Windows
                $sectionCfg['pathFromRootCompat'] = $sectionCfg['pathFromRoot'] | Remove-InvalidFileNameChars
                $sectionCfg['uri'] = $section.path # E.g. https://d.docs.live.net/0123456789abcdef/Skydrive Notebooks/mynotebook/mysectiongroup/mysection
                $sectionCfg['lastModifiedTime'] = [Datetime]::ParseExact($section.lastModifiedTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                $sectionCfg['lastModifiedTimeEpoch'] = [int][double]::Parse((Get-Date ((Get-Date $sectionCfg['lastModifiedTime']).ToUniversalTime()) -UFormat %s)) # Epoch

                $sectionCfg['pages'] = [System.Collections.ArrayList]@()

                # Build Section's pages
                if (Get-Member -InputObject $section -Name 'Page') {
                    foreach ($page in $section.Page) {
                        "$( '#' * ($LevelsFromRoot + 2) ) Building conversion configuration for $( $page.name ) [Page]" | Write-Host -ForegroundColor DarkGray

                        $previousPage = if ($sectionCfg['pages'].Count -gt 0) { $sectionCfg['pages'][$sectionCfg['pages'].Count - 1] } else { $null }
                        $pageCfg = [ordered]@{}
                        $pageCfg['notebookName'] = $cfg['notebookName']
                        $pageCfg['notesBaseDirectory'] = $cfg['notesBaseDirectory']
                        $pageCfg['notesDirectory'] = $cfg['notesDirectory']
                        $pageCfg['sectionGroupUri'] = $cfg['uri'] # Keep a reference to mt Section Group Configuration object's uri
                        $pageCfg['sectionGroupName'] = $cfg['object'].name
                        $pageCfg['sectionUri'] = $sectionCfg['uri'] # Keep a reference to my Section Configuration object's uri
                        $pageCfg['sectionName'] = $sectionCfg['object'].name
                        $pageCfg['object'] = $page # Keep a reference to my Page object
                        $pageCfg['kind'] = 'Page'
                        $pageCfg['id'] = $page.ID # E.g. {3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096991}
                        $pageCfg['nameCompat'] = $page.name | Remove-InvalidFileNameChars
                        $pageCfg['levelsFromRoot'] = $sectionCfg['levelsFromRoot']
                        $pageCfg['pathFromRoot'] = "$( $sectionCfg['pathFromRoot'] )$( [io.path]::DirectorySeparatorChar )$( $pageCfg['nameCompat'] )"
                        $pageCfg['pathFromRootCompat'] = $pageCfg['pathFromRoot'] | Remove-InvalidFileNameChars
                        $pageCfg['uri'] = "$( $sectionCfg['object'].path )/$( $page.name )" # There's no $page.path property, so we generate one. E.g. https://d.docs.live.net/0123456789abcdef/Skydrive Notebooks/mynotebook/mysectiongroup/mysection/mypage
                        $pageCfg['dateTime'] = [Datetime]::ParseExact($page.dateTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                        $pageCfg['lastModifiedTime'] = [Datetime]::ParseExact($page.lastModifiedTime, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                        $pageCfg['lastModifiedTimeEpoch'] = [int][double]::Parse((Get-Date ((Get-Date $pageCfg['lastModifiedTime']).ToUniversalTime()) -UFormat %s)) # Epoch
                        $pageCfg['pageLevel'] = $page.pageLevel -as [int]
                        $pageCfg['conversion'] = $config['conversion']['value']
                        $pageCfg['pagePrefix'] = & {
                            # 9 different scenarios
                            if ($pageCfg['pageLevel'] -eq 1) {
                                # 1 -> 1, 2 -> 1, or 3 -> 1
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
                                        # 3 -> 2 (or 4 -> 2, but 4th level subpages don't exist, but technically this supports it)
                                        $split = $previousPage['filePathRel'].Split([io.path]::DirectorySeparatorChar)
                                        $index = $pageCfg['pageLevel'] - 1 - 1 # If page level n, the prefix should be n-1
                                        if ($index -lt 0) {
                                            $index = 0 # The shallowest subpage must be a child of a first level page, i.e. $split[0]
                                        }
                                        "$( $split[0..$index] -join [io.path]::DirectorySeparatorChar )$( [io.path]::DirectorySeparatorChar )"
                                    }
                                }else {
                                    '' # Should never end up here
                                }
                            }
                        }
                        # Win32 path limits. E.g. 'C:\path\to\file' or 'C:\path\to\folder'
                        #   Absolute path:
                        #   - Win32: Max 259 characters for files, Max 247 characters for directories.
                        #   File or directory name:
                        #   - Max 255 characters long for file or folder names
                        # Non-Win32 path limits. E.g. '\\?\C:\path\to\file' or '\\?\C:\path\to\folder'. Prefixing with '\\?\' allows Windows Powershell <= 5 (based on Win32) to support long absolute paths.
                        #   Absolute path:
                        #   - N.A.
                        #   File or directory name:
                        #   - Max 255 characters long for file or folder names
                        # See: https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file?redirectedfrom=MSDN#maxpath

                        # Normalize the final .md file path. Page names can be very long, and can exceed the max absolute path length, or max file or folder name on a Windows system.
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
                            $filePathRel | Truncate-PathFileName -Length $config['mdFileNameAndFolderNameMaxLength']['value'] # Truncate to no more than 255 characters so we don't hit the folder name limit on most file systems on Windows / Linux
                        }
                        $pageCfg['filePathRelUnderscore'] = $pageCfg['filePathRel'].Replace( [io.path]::DirectorySeparatorChar, '_' )
                        $pageCfg['filePathNormal'] = & {
                            $pathWithoutExtension = if ($config['prefixFolders']['value'] -eq 2) {
                                [io.path]::combine( $cfg['notesDirectory'], $sectionCfg['nameCompat'], "$( $pageCfg['filePathRelUnderscore'] )" )
                            }else {
                                [io.path]::combine( $cfg['notesDirectory'], $sectionCfg['nameCompat'], "$( $pageCfg['filePathRel'] )" )
                            }
                            "$( $pathWithoutExtension | Truncate-PathFileName -Length ($config['mdFileNameAndFolderNameMaxLength']['value'] - 3) ).md" # Truncate to no more than 255 characters so we don't hit the file name limit on Windows / Linux
                        }
                        $pageCfg['filePathLong'] = "\\?\$( $pageCfg['filePathNormal'] )" # A non-Win32 path. Prefixing with '\\?\' allows Windows Powershell <= 5 (based on Win32) to support long absolute paths.
                        $pageCfg['filePath'] = if ($PSVersionTable.PSVersion.Major -le 5) {
                            $pageCfg['filePathLong'] # Add support for long paths on Powershell 5
                        }else {
                            $pageCfg['filePathNormal'] # Powershell Core supports long file paths
                        }
                        $pageCfg['fileDirectory'] = Split-Path $pageCfg['filePathNormal'] -Parent
                        $pageCfg['fileName'] = Split-Path $pageCfg['filePathNormal'] -Leaf
                        $pageCfg['fileExtension'] = if ($pageCfg['filePathNormal'] -match '(\.[^.]+)$') { $matches[1] } else { '' }
                        $pageCfg['fileBaseName'] = $pageCfg['fileName'] -replace "$( [regex]::Escape($pageCfg['fileExtension']) )$", ''
                        $pageCfg['pdfExportFilePathTmp'] = [io.path]::combine( (Split-Path $pageCfg['filePath'] -Parent ), "$( $pageCfg['id'] )-$( $pageCfg['lastModifiedTimeEpoch'] ).pdf" ) # Publishing a .pdf seems to be limited to 204 characters. So we will export the .pdf to a unique file name, then rename it to the actual name
                        $pageCfg['pdfExportFilePath'] = if ( ($pageCfg['fileName'].Length + ('.pdf'.Length - '.md'.Length)) -le $config['mdFileNameAndFolderNameMaxLength']['value']) {
                            $pageCfg['filePath'] -replace '\.md$', '.pdf'
                        }else {
                            $pageCfg['filePath'] -replace '.\.md$', '.pdf' # Trim 1 character in the basename when replacing the extension
                        }
                        $pageCfg['levelsPrefix'] = if ($config['medialocation']['value'] -eq 2) {
                            ''
                        }else {
                            if ($config['prefixFolders']['value'] -eq 2) {
                                "$( '../' * ($pageCfg['levelsFromRoot'] + 1 - 1) )"
                            }else {
                                "$( '../' * ($pageCfg['levelsFromRoot'] + $pageCfg['pageLevel'] - 1) )"
                            }
                        }
                        $pageCfg['tmpPath'] = & {
                            $dateNs = Get-Date -Format "yyyy-MM-dd-HH-mm-ss-fffffff"
                            if ($env:OS -match 'windows') {
                                # Ensure $env:TEMP is not MSDOS 8.3 shortened name, but the actual full path. See: https://superuser.com/questions/1524767/powershell-uses-the-short-8-3-form-for-envtemp
                                [io.path]::combine((Get-Item $env:TEMP).FullName, $cfg['notebookName'], $dateNs)
                            }else {
                                [io.path]::combine('/tmp', $cfg['notebookName'], $dateNs)
                            }
                        }
                        $pageCfg['mediaParentPath'] = if ($config['medialocation']['value'] -eq 2) {
                            $pageCfg['fileDirectory']
                        }else {
                            $cfg['notesBaseDirectory']
                        }
                        $pageCfg['mediaPath'] = [io.path]::combine( $pageCfg['mediaParentPath'], 'media' )
                        $pageCfg['mediaParentPathPandoc'] = [io.path]::combine( $pageCfg['tmpPath'] ).Replace( [io.path]::DirectorySeparatorChar, '/' ) # Pandoc outputs paths in markdown with with front slahes after the supplied <mediaPath>, e.g. '<mediaPath>/media/image.png'. So let's use a front-slashed supplied mediaPath
                        $pageCfg['mediaPathPandoc'] = [io.path]::combine( $pageCfg['tmpPath'], 'media').Replace( [io.path]::DirectorySeparatorChar, '/' ) # Pandoc outputs paths in markdown with with front slahes after the supplied <mediaPath>, e.g. '<mediaPath>/media/image.png'. So let's use a front-slashed supplied mediaPath
                        $pageCfg['docxExportFilePath'] = if ($config['docxNamingConvention']['value'] -eq 1) {
                            [io.path]::combine( $cfg['notesDocxDirectory'], "$( $pageCfg['id'] )-$( $pageCfg['lastModifiedTimeEpoch'] ).docx" )
                        }else {
                            [io.path]::combine( $cfg['notesDocxDirectory'], "$( $pageCfg['pathFromRootCompat'] ).docx" )
                        }
                        $pagexml = Get-OneNotePageContent -OneNoteConnection $OneNoteConnection -PageId $pageCfg['object'].ID
                        $ns = new-object Xml.XmlNamespaceManager $pagexml.NameTable
                        $ns.AddNamespace("one", $pagexml.DocumentElement.NamespaceURI)
                        $pageCfg['insertedAttachments'] = @(
                            & {
                                # Search recursively for all attachment(s). This includes attachments nested in tables etc.
                                $insertedFiles = $pagexml.SelectNodes("//one:InsertedFile", $ns)
                                foreach ($i in $insertedFiles) {
                                    $attachmentCfg = [ordered]@{}
                                    $attachmentCfg['object'] =  $i
                                    $attachmentCfg['nameCompat'] =  $i.preferredName | Remove-InvalidFileNameCharsInsertedFiles
                                    $attachmentCfg['markdownFileName'] =  $attachmentCfg['nameCompat'] | Encode-Markdown -Uri
                                    $attachmentCfg['source'] =  $i.pathCache
                                    $attachmentCfg['destination'] =  [io.path]::combine( $pageCfg['mediaPath'], $attachmentCfg['nameCompat'] )

                                    $attachmentCfg
                                }
                            }
                        )
                        $pageCfg['inkDrawings'] = @(
                            if ($config['mergeInkStrokes']['value'] -eq 2) {
                                # Pen/hand written drawings at the page level. Each drawing is a single stroke, exported as a separate image in the .docx in this order
                                foreach ($i in $pagexml.SelectNodes("/one:Page/one:InkDrawing[one:Position and one:Size and one:Data]", $ns)) {
                                    $inkDrawingCfg = [ordered]@{}
                                    $inkDrawingCfg['x'] = [double]$i.Position.x # Points
                                    $inkDrawingCfg['y'] = [double]$i.Position.y # Points
                                    $inkDrawingCfg['width'] = [double]$i.Size.width # Points
                                    $inkDrawingCfg['height'] = [double]$i.Size.height # Points
                                    $inkDrawingCfg['data'] = $i.Data # Base64-encoded Ink Serialized Format (ISF)

                                    $inkDrawingCfg
                                }
                            }
                        )
                        $pageCfg['images'] = @(
                            if ($pageCfg['inkDrawings'].Count -gt 0) {
                                # Images that may have been drawn on
                                Get-OneNotePageImages -PageXml $pagexml
                            }
                        )
                        $pageCfg['mutations'] = @(
                            # Markdown mutations. Each search and replace is done against a string containing the entire markdown content

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
                                        searchRegex = '^\s*'
                                        replacement = & {
                                            $heading = "# $( $pageCfg['object'].name )"
                                            if ($config['headerTimestampEnabled']['value'] -eq 1) {
                                                $heading += "`n`nCreated: $(  $pageCfg['dateTime'].ToString('yyyy-MM-dd HH:mm:ss zz00') )"
                                                $heading += "`n`nModified: $(  $pageCfg['lastModifiedTime'].ToString('yyyy-MM-dd HH:mm:ss zz00') )"
                                                $heading += "`n`n---`n`n"
                                            }
                                            $heading
                                        }
                                    }
                                )
                            }
                            if ($config['keepspaces']['value'] -eq 1 ) {
                                @{
                                    description = 'Clear extra newlines between unordered (bullet) and ordered (numbered) list items, non-breaking spaces from blank lines, and `>` after unordered lists'
                                    replacements = @(
                                        # Remove non-breaking spaces
                                        @{
                                            searchRegex = [regex]::Escape([char]0x00A0)
                                            replacement = ''
                                        }
                                        # Remove an extra newline between each occurrence of '- some unordered list item'
                                        @{
                                            searchRegex = '(\s*)- ([^\r\n]*)\r*\n\r*\n(?=\s*-)'
                                            replacement = "`$1- `$2`n"
                                        }
                                        # Remove an extra newline between each occurrence of '1. some ordered list item'
                                        @{
                                            searchRegex = '(\s*)(\d+\.) ([^\r\n]*)\r*\n\r*\n(?=\s*\d+\.)'
                                            replacement = "`$1`$2 `$3`n"
                                        }
                                        # Remove all '>' occurrences immediately following unordered lists
                                        @{
                                            searchRegex = '\n>[ ]*'
                                            replacement = "`n"
                                        }
                                    )
                                }
                            }
                            if ($config['keepescape']['value'] -eq 1) {
                                @{
                                    description = "Clear all '\' characters"
                                    replacements = @(
                                        @{
                                            searchRegex = [regex]::Escape('\')
                                            replacement = ''
                                        }
                                    )
                                }
                            }
                            elseif ($config['keepescape']['value'] -eq 2) {
                                @{
                                    description = "Clear all '\' characters except those preceding alphanumeric characters"
                                    replacements = @(
                                        @{
                                            searchRegex = '\\([^A-Za-z0-9])'
                                            replacement = '$1'
                                        }
                                    )
                                }
                            }
                            & {
                                if ($config['newlineCharacter']['value'] -eq 1) {
                                    @{
                                        description = "Use LF for newlines"
                                        replacements = @(
                                            @{
                                                searchRegex = '\r*\n'
                                                replacement = "`n"
                                            }
                                        )
                                    }
                                }else {
                                    @{
                                        description = "Use CRLF for newlines"
                                        replacements = @(
                                            @{
                                                searchRegex = '\r*\n'
                                                replacement = "`r`n"
                                            }
                                        )
                                    }
                                }
                            }
                        )
                        $pageCfg['directoriesToCreate'] = @(
                            # The directories to be created. These directories should never hit the absolute path, file name, or directory name limits on Windows
                            @(
                                $cfg['notesDocxDirectory']
                                $cfg['notesDirectory']
                                $pageCfg['tmpPath']
                                $pageCfg['fileDirectory']
                                $pageCfg['mediaPath']
                            ) | Select-Object -Unique
                        )
                        $pageCfg['directoriesToDelete'] = @(
                            $pageCfg['tmpPath']
                        )
                        $pageCfg['directorySeparatorChar'] = [io.path]::DirectorySeparatorChar

                        # Populate the pages array (needed even when -AsArray switch is not on, because we need this section's pages' state to know whether there are duplicate page names)
                        $sectionCfg['pages'].Add( $pageCfg ) > $null

                        if (!$AsArray) {
                            # Send the configuration immediately down the pipeline
                            $pageCfg
                        }
                    }
                }else {
                    "Ignoring empty Section: $( $sectionCfg['pathFromRoot'] )" | Write-Host -ForegroundColor DarkGray
                }

                # Populate the sections array
                if ($AsArray) {
                    $cfg['sections'].Add( $sectionCfg ) > $null
                }
            }
        }

        $cfg['sectionGroups'] = [System.Collections.ArrayList]@()

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
                    if (!$config['dryRun']['value']) {
                        $item = New-Item -Path $d -ItemType Directory -Force -ErrorAction Stop
                    }
                }catch {
                    Write-Error "Failed to create directory $d" -ErrorAction Continue
                    throw
                }
            }

            if ($config['usedocx']['value'] -eq 1) {
                # Remove any existing docx files, don't proceed if it fails
                try {
                    "Removing existing docx file: $( $pageCfg['docxExportFilePath'] )" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        if (Test-Path -LiteralPath $pageCfg['docxExportFilePath']) {
                            Remove-Item -LiteralPath $pageCfg['docxExportFilePath'] -Force -ErrorAction Stop
                        }
                    }
                }catch {
                    Write-Error "Failed to remove existing docx file $( $pageCfg['docxExportFilePath'] )" -ErrorAction Continue
                    throw
                }
            }

            # Publish OneNote page to Word, don't proceed if it fails
            if (! (Test-Path -LiteralPath $pageCfg['docxExportFilePath']) ) {
                try {
                    "Publishing new docx file: $( $pageCfg['docxExportFilePath'] )" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        Publish-OneNotePage -OneNoteConnection $OneNoteConnection -PageId $pageCfg['object'].ID -Destination $pageCfg['docxExportFilePath'] -PublishFormat 'pfWord'
                    }
                }catch {
                    Write-Error "Failed to publish page to docx file $( $pageCfg['docxExportFilePath'] )" -ErrorAction Continue
                    throw
                }
            }else {
                "Existing docx file: $( $pageCfg['docxExportFilePath'] )" | Write-Verbose
            }

            # Publish OneNote page to pdf, don't proceed if it fails
            if ($config['exportPdf']['value'] -eq 2) {
                if (! (Test-Path -LiteralPath $pageCfg['pdfExportFilePath']) ) {
                    try {
                        "Publishing new pdf file: $( $pageCfg['pdfExportFilePath'] )" | Write-Verbose
                        if (!$config['dryRun']['value']) {
                            Publish-OneNotePage -OneNoteConnection $OneNoteConnection -PageId $pageCfg['object'].ID -Destination $pageCfg['pdfExportFilePathTmp'] -PublishFormat 'pfPdf'
                            Move-Item $pageCfg['pdfExportFilePathTmp'] $pageCfg['pdfExportFilePath']
                        }
                        "pdf file ready: $( $pageCfg['pdfExportFilePath'] )" | Write-Host -ForegroundColor Green
                    }catch {
                        Write-Error "Failed to publish page to pdf file $( $pageCfg['pdfExportFilePath'] )" -ErrorAction Continue
                        throw
                    }
                }else {
                    "Existing pdf file: $( $pageCfg['pdfExportFilePath'] )" | Write-Host -ForegroundColor Green
                }
            }

            # https://gist.github.com/heardk/ded40b72056cee33abb18f3724e0a580
            # Convert .docx to .md, don't proceed if it fails
            $stderrFile = "$( $pageCfg['tmpPath'] )/pandoc-stderr.txt"
            try {
                # Start-Process has no way of capturing stderr / stdterr to variables, so we need to use temp files.
                "Converting docx file to markdown file: $( $pageCfg['filePath'] )" | Write-Verbose
                if (!$config['dryRun']['value']) {
                    $argumentList = @(
                        '-f'
                        'docx'
                        '-t'
                        $pageCfg['conversion']
                        '-i'
                        if ($pageCfg['docxExportFilePath'] -match ' ') {
                            "`"$( $pageCfg['docxExportFilePath'] )`"" # Add double-quotes to path containing spaces
                        }else {
                            $pageCfg['docxExportFilePath']
                        }
                        '-o'
                        if ($pageCfg['filePathNormal'] -match ' ') {
                            "`"$( $pageCfg['filePathNormal'] )`"" # Add double-quotes to path containing spaces
                        }else {
                            $pageCfg['filePathNormal']
                        }
                        '--wrap=none'
                        '--markdown-headings=atx'
                        if ($pageCfg['mediaParentPathPandoc'] -match ' ') {
                            "`"--extract-media=$( $pageCfg['mediaParentPathPandoc'] )`"" # Add double-quotes to path containing spaces
                        }else {
                            "--extract-media=$( $pageCfg['mediaParentPathPandoc'] )"
                        }
                    )
                    "Command line: pandoc.exe $argumentList" | Write-Verbose
                    $process = Start-Process -ErrorAction Stop -RedirectStandardError $stderrFile -PassThru -NoNewWindow -Wait -FilePath pandoc.exe -ArgumentList $argumentList # extracts into ./media of the supplied folder
                    if ($process.ExitCode -ne 0) {
                        $stderr = Get-Content $stderrFile -Raw
                        throw "pandoc error: $stderr"
                    }
                }
            }catch {
                Write-Error "Failed to convert docx file $( $pageCfg['docxExportFilePath'] ) to markdown file $( $pageCfg['filePathNormal'] )"
                throw
            }finally {
                if (Test-Path $stderrFile) {
                    Remove-Item $stderrFile -Force
                }
            }

            # Merge the strokes of pen/hand written drawings into single images
            if ($pageCfg['inkDrawings'].Count -gt 0) {
                try {
                    "Merging the strokes of $( $pageCfg['inkDrawings'].Count ) ink drawing(s) into single images" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        Merge-InkDrawingStrokes -PageConfig $pageCfg
                    }
                }catch {
                    Write-Error "Failed to merge ink drawing strokes of $( $pageCfg['filePathNormal'] ). Exception: $( $_.Exception.Message )" -ErrorAction Continue
                }
            }

            # Cleanup Word files
            if ($config['keepdocx']['value'] -eq 1) {
                try {
                    "Removing existing docx file: $( $pageCfg['docxExportFilePath'] )" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        if (Test-Path -LiteralPath $pageCfg['docxExportFilePath']) {
                            Remove-Item -LiteralPath $pageCfg['docxExportFilePath'] -Force -ErrorAction Stop
                        }
                    }
                }catch {
                    Write-Error "Failed to remove existing docx file $( $pageCfg['docxExportFilePath'] ). Exception: $( $_.Exception.Message )" -ErrorAction Continue
                }
            }

            # Save any attachments
            foreach ($attachmentCfg in $pageCfg['insertedAttachments']) {
                try {
                    "Saving inserted attachment: $( $attachmentCfg['destination'] )" | Write-Verbose
                    if (!$config['dryRun']['value']) {
                        Copy-Item -Path $attachmentCfg['source'] -Destination $attachmentCfg['destination'] -Force -ErrorAction Stop
                    }
                }catch {
                    Write-Error "Failed to save attachment from $( $attachmentCfg['source'] ) to $( $attachmentCfg['destination'] ). Exception: $( $_.Exception.Message )" -ErrorAction Continue
                }
            }

            # Rename images to have unique names - NoteName-Image#-HHmmssff.xyz
            if (!$config['dryRun']['value']) {
                $images = Get-ChildItem -Path $pageCfg['mediaPathPandoc'] -Recurse -Force -ErrorAction SilentlyContinue
                foreach ($image in $images) {
                    # Rename Image
                    try {
                        $newimageName = if ($config['medialocation']['value'] -eq 2) {
                            "$( $pageCfg['filePathRelUnderscore'] )-$($image.BaseName)$($image.Extension)"
                        }else {
                            "$( $pageCfg['pathFromRootCompat'] )-$($image.BaseName)$($image.Extension)"
                        }
                        $newimagePath = [io.path]::combine( $pageCfg['mediaPath'], $newimageName )
                        "Moving image: $( $image.FullName ) to $( $newimagePath )" | Write-Verbose
                        if (!$config['dryRun']['value']) {
                            $item = Move-Item -Path "$( $image.FullName )" -Destination $newimagePath -Force -ErrorAction Stop -PassThru
                        }
                    }catch {
                        Write-Error "Failed to rename image $( $image.FullName ) to $( $item.FullName ). Exception: $( $_.Exception.Message )" -ErrorAction Continue
                    }
                    # Mutate markdown content with new image references
                    try {
                        "Mutation of markdown: Rename image references. Find: '$( $image.Name )', Replacement: '$( $newimageName )'" | Write-Verbose
                        if (!$config['dryRun']['value']) {
                            $content = Get-Content -LiteralPath $pageCfg['filePath'] -Raw -ErrorAction Stop # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
                            $content = $content.Replace("$($image.Name)", "$($newimageName)")
                            Set-ContentNoBom -LiteralPath $pageCfg['filePath'] -Value $content -ErrorAction Stop # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
                        }
                    }catch {
                        Write-Error "Failed to rename image references to $( $newimageName ). Exception: $( $_.Exception.Message )" -ErrorAction Continue
                    }
                }
            }

            # Mutate markdown content
            try {
                if (!$config['dryRun']['value']) {
                    # Get markdown content
                    $content = @( Get-Content -LiteralPath $pageCfg['filePath'] -ErrorAction Stop ) # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
                    $content = @(
                        if ($content.Count -gt 6) {
                            # Discard first 6 lines which contain a header, created date, and time. We are going to add our own header
                            $content[6..($content.Count - 1)]
                        }else {
                            # Empty page
                            ''
                        }
                    ) -join "`n"
                }

                # Mutate
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        try {
                            "Mutation of markdown: $( $m['description'] ). Regex: '$( $r['searchRegex'] )', Replacement: '$( $r['replacement'].Replace("`r", '\r').Replace("`n", '\n') )'" | Write-Verbose
                            if (!$config['dryRun']['value']) {
                                $content = $content -replace $r['searchRegex'], $r['replacement']
                            }
                        }catch {
                            Write-Error "Failed to mutate markdown content with mutation '$( $m['description'] )'. Exception: $( $_.Exception.Message )"
                        }
                    }
                }
                if (!$config['dryRun']['value']) {
                    Set-ContentNoBom -LiteralPath $pageCfg['filePath'] -Value $content -ErrorAction Stop # Use -LiteralPath so that characters like '(', ')', '[', ']', '`', "'", '"' are supported. Or else we will get an error "Cannot find path 'xxx' because it does not exist"
                }
            }catch {
                Write-Error "Failed to mutate markdown content: $( $_.Exception.Message )"
            }

            "Markdown file ready: $( $pageCfg['filePathNormal'] )" | Write-Host -ForegroundColor Green
        }catch {
            Write-Error "Failed to convert page: $( $pageCfg['pathFromRoot'] )" -ErrorAction Continue
            if ($ErrorActionPreference -eq 'Stop') {
                throw
            }else {
                Write-Error -Message $_.Exception.Message
                Write-Error -Message $_.ScriptStackTrace
            }
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
        $ErrorCollection | Where-Object { (Get-Member -InputObject $_ -Name 'CategoryInfo') -and ($_.CategoryInfo.Reason -match 'WriteErrorException') } | Write-Host -ForegroundColor Red
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
                $hierarchy.Notebooks.Notebook | Where-Object { $_.Name -eq $config['targetNotebook']['value'] }
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

        "`nNotebooks to convert:" | Write-Host -ForegroundColor Cyan
        $notebooks.name | Write-Host -ForegroundColor Green

        # Convert the notebook(s)
        $pageConversionConfigsAll = @()
        foreach ($notebook in $notebooks) {
            "`nConverting notebook '$( $notebook.name )'... (Ignoring deleted notes)" | Write-Host -ForegroundColor Cyan
            New-SectionGroupConversionConfig -OneNoteConnection $OneNote -NotesDestination $config['notesdestpath']['value'] -Config $config -SectionGroups $notebook -LevelsFromRoot 0 -ErrorVariable +totalerr | Tee-Object -Variable pageConversionConfigs | Convert-OneNotePage -OneNoteConnection $OneNote -Config $config -ErrorVariable +totalerr
            "`nDone converting notebook '$( $notebook.name )' with $( ($pageConversionConfigs | Measure-object).Count ) notes." | Write-Host -ForegroundColor Cyan
            $pageConversionConfigsAll += $pageConversionConfigs
        }

        # Export all Page Conversion Configuration objects as .json, which is useful for debugging
        if ($ConversionConfigurationExportPath) {
            "`nExporting Page Conversion Configuration as JSON file with $( $pageConversionConfigsAll.Count ) objects: $ConversionConfigurationExportPath" | Write-Host -ForegroundColor Cyan
            $pageConversionConfigsAll | ConvertTo-Json -Depth 100 | Out-File $ConversionConfigurationExportPath -Encoding utf8 -Force
        }
    }catch {
        if ($ErrorActionPreference -eq 'Stop') {
            throw
        }else {
            Write-Error -Message $_.Exception.Message
            Write-Error -Message $_.ScriptStackTrace
            "If you're seeing errors on the first run, check the FAQ in README.md to resolve them: " | Write-Host -ForegroundColor Yellow
            "  https://github.com/theohbrothers/ConvertOneNote2MarkDown#faq" | Write-Host -ForegroundColor Yellow
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
