Function Remove-InvalidFileNameChars {
    param(
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]$Name,
        [switch]$KeepPathSpaces
    )

    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    $newName = $newName -replace "\[", "("
    $newName = $newName -replace "\]", ")"
    $newName =  if ($KeepPathSpaces) {
                    $newName -replace "\s", " "
                } else {
                    $newName -replace "\s", "-"
                }
    $newName = $newName.Substring(0, $(@{$true = 130; $false = $newName.length }[$newName.length -gt 150]))
    return $newName.Trim() # Remove boundary whitespaces
}
Function Remove-InvalidFileNameCharsInsertedFiles {
    param(
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true)]
        [string]$Name,
        [string]$Replacement = "",
        [string]$SpecialChars = "#$%^*[]'<>!@{};",
        [switch]$KeepPathSpaces
    )

    $rePattern = ($SpecialChars.ToCharArray() | ForEach-Object { [regex]::Escape($_) }) -join "|"

    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    $newName = $newName -replace $rePattern, ""
    $newName =  if ($KeepPathSpaces) {
                    $newName -replace "\s", " "
                } else {
                    $newName -replace "\s", "-"
                }
    return $newName.Trim() # Remove boundary whitespaces
}

Function ProcessSections ($group, $FilePath) {
    foreach ($section in $group.Section) {
        "--------------" | Write-Host
        "### " + $section.Name | Write-Host
        $sectionFileName = "$($section.Name)" | Remove-InvalidFileNameChars -KeepPathSpaces:($keepPathSpaces -eq 2)
        $item = New-Item -Path "$($FilePath)" -Name "$($sectionFileName)" -ItemType "directory" -Force -ErrorAction SilentlyContinue
        "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host
        [int]$previouspagelevel = 1
        [string]$previouspagenamelevel1 = ""
        [string]$previouspagenamelevel2 = ""
        [string]$pageprefix = ""

        foreach ($page in $section.Page) {
            # set page variables
            $recurrence = 1
            $pagelevel = $page.pagelevel
            $pagelevel = $pagelevel -as [int]
            $pageid = ""
            $pageid = $page.ID
            $pagename = ""
            $pagename = $page.name | Remove-InvalidFileNameChars -KeepPathSpaces:($keepPathSpaces -eq 2)
            $fullexportdirpath = ""
            $fullexportdirpath = "$($FilePath)\$($sectionFileName)"
            $fullfilepathwithoutextension = ""
            $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
            $fullexportpath = ""
            #$fullexportpath = "$($fullfilepathwithoutextension).docx"


            # process for subpage prefixes
            if ($pagelevel -eq 1) {
                $pageprefix = ""
                $previouspagenamelevel1 = $pagename
                $previouspagenamelevel2 = ""
                $previouspagelevel = 1
                "#### " + $page.name | Write-Host
            }
            elseif ($pagelevel -eq 2) {
                $pageprefix = "$($previouspagenamelevel1)"
                $previouspagenamelevel2 = $pagename
                $previouspagelevel = 2
                "##### " + $page.name | Write-Host
            }
            elseif ($pagelevel -eq 3) {
                if ($previouspagelevel -eq 2) {
                    $pageprefix = "$($previouspagenamelevel1)$($prefixjoiner)$($previouspagenamelevel2)"
                }
                # level 3 under level 1, without a level 2
                elseif ($previouspagelevel -eq 1) {
                    $pageprefix = "$($previouspagenamelevel1)$($prefixjoiner)"
                }
                #and if previous was 3, do nothing/keep previous label
                $previouspagelevel = 3
                "####### " + $page.name | Write-Host
            }

            #if level 2 or 3 (i.e. has a non-blank pageprefix)
            if ($pageprefix) {
                #create filename prefixes and filepath if prefixes selected
                if ($prefixFolders -eq 2) {
                    $pagename = "$($pageprefix)_$($pagename)"
                    $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
                }
                #all else/default, create subfolders and filepath if subfolders selected
                else {
                    $item = New-Item -Path "$($fullexportdirpath)\$($pageprefix)" -ItemType "directory" -Force -ErrorAction SilentlyContinue
                    "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host | Out-Null
                    $fullexportdirpath = "$($fullexportdirpath)\$($pageprefix)"
                    $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
                    $levelsprefix = "../" * ($levelsfromroot + $pagelevel - 1) + ".."
                }
            }
            else {
                $levelsprefix = "../" * ($levelsfromroot) + ".."
            }

            # set media location (central media folder at notebook-level or adjacent to .md file) based on initial user prompt
            if ($medialocation -eq 2) {
                $mediaPath = $fullexportdirpath.Replace('\', '/') # Normalize markdown media paths to use front slashes, i.e. '/'
                $levelsprefix = ""
            }
            else {
                $mediaPath = $NotebookFilePath.Replace('\', '/') # Normalize markdown media paths to use front slashes, i.e. '/'
            }
            $mediaPath = $mediaPath.Substring(0, 1).tolower() + $mediaPath.Substring(1) # Normalize markdown media paths to use a lowercased drive letter

            # in case multiple pages with the same name exist in a section, postfix the filename. Run after pages
            if ([System.IO.File]::Exists("$($fullfilepathwithoutextension).md")) {
                #continue
                $pagename = "$($pagename)-$recurrence"
                $recurrence++
            }

            $fullexportpath = "$($NotebookFilePath)\docx\$($pagename).docx"

            # use existing or create new docx files
            if ($usedocx -eq 2) {
                # Only create new docx if doesn't exist
                if (![System.IO.File]::Exists($fullexportpath)) {
                    # publish OneNote page to Word
                    try {
                        $OneNote.Publish($pageid, $fullexportpath, "pfWord", "")
                    }
                    catch {
                        Write-Host "Error while publishing file '$($page.name)' to docx: $($Error[0].ToString())" -ForegroundColor Red
                        $totalerr += "Error while publishing file '$($page.name)' to docx: $($Error[0].ToString())`r`n"
                    }
                }
            }
            else {
                # remove any existing docx files
                if ([System.IO.File]::Exists($fullexportpath)) {
                    try {
                        Remove-Item -path $fullexportpath -Force -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-Host "Error removing intermediary '$($page.name)' docx file: $($Error[0].ToString())" -ForegroundColor Red
                        $totalerr += "Error removing intermediary '$($page.name)' docx file: $($Error[0].ToString())`r`n"
                    }
                }

                # publish OneNote page to Word
                try {
                    $OneNote.Publish($pageid, $fullexportpath, "pfWord", "")
                }
                catch {
                    Write-Host "Error while publishing file '$($page.name)' to docx: $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while publishing file '$($page.name)' to docx: $($Error[0].ToString())`r`n"
                }
            }

            # https://gist.github.com/heardk/ded40b72056cee33abb18f3724e0a580
            try {
                pandoc.exe -f  docx -t $converter-simple_tables-multiline_tables-grid_tables+pipe_tables -i $fullexportpath -o "$($fullfilepathwithoutextension).md" --wrap=none --markdown-headings=atx --extract-media="$($mediaPath)"
            }
            catch {
                Write-Host "Error while converting file '$($page.name)' to md: $($Error[0].ToString())" -ForegroundColor Red
                $totalerr += "Error while converting file '$($page.name)' to md: $($Error[0].ToString())`r`n"
            }

            # export inserted file objects, removing any escaped symbols from filename so that links to them actually work
            [xml]$pagexml = ""
            $OneNote.GetPageContent($pageid, [ref]$pagexml, 7)
            $pageinsertedfiles = $pagexml.Page.Outline.OEChildren.OE | Where-Object { $_.InsertedFile }
            foreach ($pageinsertedfile in $pageinsertedfiles) {
                $item = New-Item -Path "$($mediaPath)" -Name "media" -ItemType "directory" -Force -ErrorAction SilentlyContinue
                "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host | Out-Null
                $destfilename = ""
                try {
                    $destfilename = $pageinsertedfile.InsertedFile.preferredName | Remove-InvalidFileNameCharsInsertedFiles -KeepPathSpaces:($keepPathSpaces -eq 2)
                    Copy-Item -Path "$($pageinsertedfile.InsertedFile.pathCache)" -Destination "$($mediaPath)\media\$($destfilename)" -Force
                }
                catch {
                    Write-Host "Error while copying file object '$($pageinsertedfile.InsertedFile.preferredName)' for page '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while copying file object '$($pageinsertedfile.InsertedFile.preferredName)' for page '$($page.name)': $($Error[0].ToString())`r`n"
                }
                # Change MD file Object Name References
                try {
                    $pageinsertedfile2 = $pageinsertedfile.InsertedFile.preferredName.Replace("$", "\$").Replace("^", "\^").Replace("'", "\'")
                    ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw).Replace("$($pageinsertedfile2)", "[$($destfilename)]($($mediaPath)/media/$($destfilename))")) | Set-Content -Path "$($fullfilepathwithoutextension).md"

                }
                catch {
                    Write-Host "Error while renaming file object name references to '$($pageinsertedfile.InsertedFile.preferredName)' for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while renaming file object name references to '$($pageinsertedfile.InsertedFile.preferredName)' for file '$($page.name)': $($Error[0].ToString())`r`n"
                }
            }

            # add YAML
            $orig = @(
                Get-Content -path "$($fullfilepathwithoutextension).md"
            )
            $orig[0] = "# $($page.name)"
            if ($headerTimestampEnabled -eq 2) {
                Set-Content -Path "$($fullfilepathwithoutextension).md" -Value $orig[0..0], $orig[6..($orig.Count - 1)]
            }else {
                $insert1 = "$($page.dateTime)"
                $insert1 = [Datetime]::ParseExact($insert1, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
                $insert1 = $insert1.ToString("yyyy-MM-dd HH:mm:ss
                ")
                $insert2 = "---"
                Set-Content -Path "$($fullfilepathwithoutextension).md" -Value $orig[0..0], $insert1, $insert2, $orig[6..($orig.Count - 1)]
            }

            #Clear double spaces from bullets and nonbreaking spaces from blank lines
            if ($keepspaces -eq 2 ) {
                #do nothing
            }
            else {
                try {
                    ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw -encoding utf8).Replace([char]0x00A0, [char]0x000A).Replace([char]0x000A, [char]0x000A).Replace("`r`n`r`n", "`r`n")) | Set-Content -Path "$($fullfilepathwithoutextension).md"
                }
                catch {
                    Write-Host "Error while clearing double spaces from file '$($fullfilepathwithoutextension)' : $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while clearing double spaces from file '$($fullfilepathwithoutextension)' : $($Error[0].ToString())`r`n"
                }
            }

            # rename images to have unique names - NoteName-Image#-HHmmssff.xyz
            $timeStamp = (Get-Date -Format HHmmssff).ToString()
            $timeStamp = $timeStamp.replace(':', '')
            $images = Get-ChildItem -Path "$($mediaPath)/media" -Include "*.png", "*.gif", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name.SubString(0, 5) -match "image" }
            foreach ($image in $images) {
                $newimageName = "$($pagename.SubString(0,[math]::min(30,$pagename.length)))-$($image.BaseName)-$($timeStamp)$($image.Extension)"
                # Rename Image
                try {
                    Rename-Item -Path "$($image.FullName)" -NewName $newimageName -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host "Error while renaming image '$($image.FullName)' for page '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while renaming image '$($image.FullName)' for page '$($page.name)': $($Error[0].ToString())`r`n"
                }
                # Change MD file Image filename References
                try {
                    ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw).Replace("$($image.Name)", "$($newimageName)")) | Set-Content -Path "$($fullfilepathwithoutextension).md"
                }
                catch {
                    Write-Host "Error while renaming image file name references to '$($image.Name)' for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while renaming image file name references to '$($image.Name)' for file '$($page.name)': $($Error[0].ToString())`r`n"
                }
            }

            # change MD file Image Path References
            try {
                # Change MD file Image Path References in Markdown
                ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw).Replace("$($mediaPath)/media/", "$($levelsprefix)/media/")) | Set-Content -Path "$($fullfilepathwithoutextension).md"
                # Change MD file Image Path References in HTML
                ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw).Replace("$($mediaPath)", "$($levelsprefix)")) | Set-Content -Path "$($fullfilepathwithoutextension).md"
            }
            catch {
                Write-Host "Error while renaming image file path references for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                $totalerr += "Error while renaming image file path references for file '$($page.name)': $($Error[0].ToString())`r`n"
            }

            # Clear backslash escape symbols
            if ($keepescape -eq 2 ) {
                #do nothing
            }
            else {
                ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw).Replace("\", '')) | Set-Content -Path "$($fullfilepathwithoutextension).md"
            }

            # Cleanup Word files
            try {
                if ($keepdocx -ne 2) {
                    Remove-Item -path "$fullexportpath" -Force -ErrorAction SilentlyContinue
                }

            }
            catch {
                Write-Host "Error removing intermediary '$($page.name)' docx file: $($Error[0].ToString())" -ForegroundColor Red
                $totalerr += "Error removing intermediary '$($page.name)' docx file: $($Error[0].ToString())`r`n"
            }
        }
    }
}

""
"-----------------------------------------------"
# ask for the Notes root path
"Enter the folder path that will contain your resulting Notes structure. ex. 'c:\temp\notes'"
$notesdestpath = Read-Host -Prompt "Entry"
""
"-----------------------------------------------"

#prompt for notebook
""
"-----------------------------------------------"
"'': Convert all notebooks - Default"
"'mynotebook': Convert specific notebook named 'mynotebook'"
$targetNotebook = Read-Host -Prompt "Enter name of notebook"

#prompt to use existing word docs (90% faster)
""
"-----------------------------------------------"
"1: Create new .docx files - Default"
"2: Use existing .docx files (90% faster)"
[int] $usedocx = Read-Host -Prompt "Entry"

#prompt to discard intermediate word docs
""
"-----------------------------------------------"
"1: Discard intermediate .docx files - Default"
"2: Keep .docx files"
[int] $keepdocx = Read-Host -Prompt "Entry"
""
"-----------------------------------------------"
# prompt for prefix vs subfolders
"1: Create folders for subpages (e.g. Page\Subpage.md)- Default"
"2: Add prefixes for subpages (e.g. Page_Subpage.md)"
[Int]$prefixFolders = Read-Host -Prompt "Entry"

#prompt for media in single or multiple folders
""
"-----------------------------------------------"
"1: Images stored in single 'media' folder at Notebook-level (Default)"
"2: Separate 'media' folder for each folder in the hierarchy"
[int] $medialocation = Read-Host -Prompt "Entry"

#prompt for conversion type
""
"Select conversion type"
"-----------------------------------------------"
"1: markdown (Pandoc) - Default"
"2: commonmark (CommonMark Markdown)"
"3: gfm (GitHub-Flavored Markdown)"
"4: markdown_mmd (MultiMarkdown)"
"5: markdown_phpextra (PHP Markdown Extra)"
"6: markdown_strict (original unextended Markdown)"
[int]$conversion = Read-Host -Prompt "Entry: "

#prompt to include page timestamp and separator at the top of document
"-----------------------------------------------"
"1: Include page timestamp and separator at top of document - Default"
"2: Do not include page timestamp and separator at top of document"
[int]$headerTimestampEnabled = Read-Host -Prompt "Entry"

#prompt to clear double spaces between bullets
"-----------------------------------------------"
"1: Clear double spaces in bullets - Default"
"2: Keep double spaces"
[int] $keepspaces = Read-Host -Prompt "Entry"

# prompt to clear escape symbols from md files
"-----------------------------------------------"
"1: Clear '\' symbol escape character from files - Default"
"2: Keep '\' symbol escape"
[int] $keepescape = Read-Host -Prompt "Entry"

#prompt to replace spaces with dashes in file
"-----------------------------------------------"
"1: Replace spaces with dashes i.e. '-' in file and folder names - Default"
"2: Keep spaces in file and folder names (1 space between words, removes preceding and trailing spaces)"
[int]$keepPathSpaces = Read-Host -Prompt "Entry"

# Fix encoding problems for languages other than English
$PSDefaultParameterValues['*:Encoding'] = 'utf8'

try {
    # Validate and compile configuration
    if (! (Test-Path -Path $notesdestpath -PathType Container -ErrorAction SilentlyContinue) ) {
        throw "Notes root path does not exist: $notesdestpath"
    }
    $notesdestpath = $notesdestpath.Trim('/').Trim('\') # Normalize notes paths to have no trailing slash or backslashes

    if ($prefixFolders -eq 2) {
        $prefixFolders = 2
        $prefixjoiner = "_"
    }else {
        $prefixFolders = 1
        $prefixjoiner = "\"
    }

    if ($conversion -eq 2) { $converter = "commonmark" }
    elseif ($conversion -eq 3) { $converter = "gfm" }
    elseif ($conversion -eq 4) { $converter = "markdown_mmd" }
    elseif ($conversion -eq 5) { $converter = "markdown_phpextra" }
    elseif ($conversion -eq 6) { $converter = "markdown_strict" }
    else { $converter = "markdown" }

    # Open OneNote hierarchy
    $OneNote = New-Object -ComObject OneNote.Application
    [xml]$Hierarchy = ""
    $totalerr = ""
    $OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)

    # Validate the notebooks to convert
    $notebooks = @(
        if ($targetNotebook) {
            $Hierarchy.Notebooks.Notebook | ? { $_.Name -eq $targetNotebook }
        }else {
            $Hierarchy.Notebooks.Notebook
        }
    )
    if ($notebooks.Count -eq 0) {
        if ($targetNotebook) {
            throw "Could not find notebook of name '$targetNotebook'"
        }else {
            throw "Could not find notebooks"
        }
    }

    foreach ($notebook in $notebooks) {
        " " | Write-Host
        $notebook.Name | Write-Host
        $notebookFileName = "$($notebook.Name)" | Remove-InvalidFileNameChars -KeepPathSpaces:($keepPathSpaces -eq 2)
        $item = New-Item -Path "$($notesdestpath)\" -Name "$($notebookFileName)" -ItemType "directory" -Force -ErrorAction SilentlyContinue
        "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host
        $NotebookFilePath = "$($notesdestpath)\$($notebookFileName)"
        $levelsfromroot = 0

        $item = New-Item -Path "$($NotebookFilePath)" -Name "docx" -ItemType "directory" -Force -ErrorAction SilentlyContinue
        "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host

        "==============" | Write-Host
        #process any sections that are not in a section group
        ProcessSections $notebook $NotebookFilePath

        #start looping through any top-level section groups in the notebook
        foreach ($sectiongroup1 in $notebook.SectionGroup) {
            $levelsfromroot = 1
            if ($sectiongroup1.isRecycleBin -ne 'true') {
                "# " + $sectiongroup1.Name | Write-Host
                $sectiongroupFileName1 = "$($sectiongroup1.Name)" | Remove-InvalidFileNameChars -KeepPathSpaces:($keepPathSpaces -eq 2)
                $item = New-Item -Path "$($notesdestpath)\$($notebookFileName)" -Name "$($sectiongroupFileName1)" -ItemType "directory" -Force -ErrorAction SilentlyContinue
                "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host | Out-Null
                $sectiongroupFilePath1 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)"
                ProcessSections $sectiongroup1 $sectiongroupFilePath1

                #start looping through any 2nd level section groups within the 1st level section group
                foreach ($sectiongroup2 in $sectiongroup1.SectionGroup) {
                    $levelsfromroot = 2
                    if ($sectiongroup2.isRecycleBin -ne 'true') {
                        "## " + $sectiongroup2.Name | Write-Host
                        $sectiongroupFileName2 = "$($sectiongroup2.Name)" | Remove-InvalidFileNameChars -KeepPathSpaces:($keepPathSpaces -eq 2)
                        $item = New-Item -Path "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)" -Name "$($sectiongroupFileName2)" -ItemType "directory" -Force -ErrorAction SilentlyContinue
                        "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host
                        $sectiongroupFilePath2 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)"
                        ProcessSections $sectiongroup2 $sectiongroupFilePath2

                        #start looping through any 2nd level section groups within the 1st level section group
                        foreach ($sectiongroup3 in $sectiongroup2.SectionGroup) {
                            $levelsfromroot = 3
                            if ($sectiongroup3.isRecycleBin -ne 'true') {
                                "### " + $sectiongroup3.Name | Write-Host
                                $sectiongroupFileName3 = "$($sectiongroup3.Name)" | Remove-InvalidFileNameChars -KeepPathSpaces:($keepPathSpaces -eq 2)
                                $item = New-Item -Path "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)" -Name "$($sectiongroupFileName3)" -ItemType "directory" -Force -ErrorAction SilentlyContinue
                                "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host
                                $sectiongroupFilePath3 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)"
                                ProcessSections $sectiongroup3 $sectiongroupFilePath3

                                #start looping through any 2nd level section groups within the 1st level section group
                                foreach ($sectiongroup4 in $sectiongroup3.SectionGroup) {
                                    $levelsfromroot = 4
                                    if ($sectiongroup4.isRecycleBin -ne 'true') {
                                        "#### " + $sectiongroup4.Name | Write-Host
                                        $sectiongroupFileName4 = "$($sectiongroup4.Name)" | Remove-InvalidFileNameChars -KeepPathSpaces:($keepPathSpaces -eq 2)
                                        $item = New-Item -Path "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)" -Name "$($sectiongroupFileName4)" -ItemType "directory" -Force -ErrorAction SilentlyContinue
                                        "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host
                                        $sectiongroupFilePath4 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)\$($sectiongroupFileName4)"
                                        ProcessSections $sectiongroup4 $sectiongroupFilePath4

                                        #start looping through any 2nd level section groups within the 1st level section group
                                        foreach ($sectiongroup5 in $sectiongroup4.SectionGroup) {
                                            $levelsfromroot = 5
                                            if ($sectiongroup5.isRecycleBin -ne 'true') {
                                                "#### " + $sectiongroup5.Name | Write-Host
                                                $sectiongroupFileName5 = "$($sectiongroup5.Name)" | Remove-InvalidFileNameChars -KeepPathSpaces:($keepPathSpaces -eq 2)
                                                $item = New-Item -Path "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)\$($sectiongroupFileName4)" -Name "$($sectiongroupFileName5)" -ItemType "directory" -Force -ErrorAction SilentlyContinue
                                                "Directory: $($item.FullName)" | Format-Table | Out-String | Write-Host
                                                $sectiongroupFilePath5 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)\$($sectiongroupFileName4)\$($sectiongroupFileName5)"
                                                ProcessSections $sectiongroup5 $sectiongroupFilePath5
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}catch {
    throw
}finally {
    'Cleaning up' | Write-Host

    # Release OneNote hierarchy
    if (Get-Variable -Name OneNote -ErrorAction SilentlyContinue) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
        Remove-Variable -Name OneNote -Force
    }

    "Errors: $totalerr" | Write-Host
}
