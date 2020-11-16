#Define global variables
[boolean]$global:activateDebug = 0

[boolean]$global:activateMOCForObsidian = 1
[boolean]$global:activateSubDir = 0
[boolean]$global:activateGlobalFileName = 0

[int]$global:fileCount = 0
[int]$global:pageNameRecurrenceCount = 1

[string]$global:filePrefix = "File-"
[string]$global:SpaceChar = " "

[int]$global:parsedLinksMatched = 0
[int]$global:parsedLinksNotMatched = 0
$global:parsedLinksAmbiguousMatched = 0

$global:PageIdArray = [System.Collections.ArrayList]::new()
$global:PageValueArray = [System.Collections.ArrayList]::new()

#Functions
Function Remove-InvalidFileNameChars {
    param(
        [Parameter(Mandatory = $true,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [string]$Name
    )
    
    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    $newName = $newName.Replace('*', '-').Replace('"', '').Replace('\', '-').Replace('/', '-').Replace('<', '-').Replace('>', '-').Replace(':', '').Replace('|', '-').Replace('?', '') #Obsidian non valid file name chars
    $newName = $newName.Trim()

    if (($global:activateDebug -eq 1) -and ($newName -ne $Name))
    {
        Write-Host "$($name) `nRenamed to '[[$($newName)]]'" -ForegroundColor Blue
    }

    return (((($newName -replace "\s", "$($global:SpaceChar)") -replace "\[", "(") -replace "\]", ")").Substring(0,$(@{$true=130;$false=$newName.length}[$newName.length -gt 150])))
}

Function Remove-InvalidIdNameChars {
    param(
        [Parameter(Mandatory = $true,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [string]$Name
    )
    
    
    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    return (((($newName -replace "\s", "$($global:SpaceChar)") -replace "\{", "--") -replace "\}", "--").Substring(0,$(@{$true=130;$false=$newName.length}[$newName.length -gt 150])))
}
Function Remove-InvalidFileNameCharsInsertedFiles {
    param(
        [Parameter(Mandatory = $true,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
        [string]$Name,
        [string]$Replacement  = "",
        [string]$SpecialChars = "#$%^*[]'<>!@{};"
        
    )

    $rePattern = ($SpecialChars.ToCharArray() |ForEach-Object { [regex]::Escape($_) }) -join "|"

    $newName = $Name.Split([IO.Path]::GetInvalidFileNameChars()) -join '-'
    return ($newName -replace $rePattern,"" -replace "\s","$($global:SpaceChar)")
}
  
Function saveHyperlinkToObject($object, $name)
{
    $link = ""
    $OneNote.GetHyperLinkToObject($object.ID, "", [ref]$link)

    if ($link -and $name)
    {
        $global:PageIdArray.add($name)
        $global:PageValueArray.add($link)
    }
}
Function ProcessSections ($group, $FilePath) {
    [string]$sectionGroupValue
    
    if ($global:activateMOCForObsidian -eq 1) #Header for section group moc file
    {
        [string]$sectionGroupValue = "# $($group.Name)`n`n---"
        saveHyperlinkToObject $group ($group.Name | Remove-InvalidFileNameChars)
    }

    foreach ($section in $group.Section) {
        "--------------"
        "### " + $section.Name
        $sectionFileName = "$($section.Name)" | Remove-InvalidFileNameChars
        if($global:activateSubDir -eq 1) 
        {
            New-Item -Path "$($FilePath)" -Name "$($sectionFileName)" -ItemType "directory" -ErrorAction SilentlyContinue
        }
        [int]$previouspagelevel = 1
        [string]$previouspagenamelevel1 = ""
        [string]$previouspagenamelevel2 = ""
        [string]$pageprefix = ""
        [string]$sectionValue = ""

        if ($global:activateMOCForObsidian -eq 1) #Header for section moc file
        {
            [string]$sectionValue = "# $($section.Name)`n`n---"
        }

        foreach ($page in $section.Page) {
            # set page variables
            if ($global:activateSubDir -eq 1) #Global if no sub dirs
            {
                $global:pageNameRecurrenceCount = 1
            }

            $pagelevel = $page.pagelevel
            $pagelevel = $pagelevel -as [int]
            $pageid = ""
            $pageid = $page.ID
            $pagename = ""
            $pagename = $page.name | Remove-InvalidFileNameChars
            $fullexportdirpath = ""
            
            if($global:activateSubDir -eq 1) 
            {
                $fullexportdirpath = "$($FilePath)\$($sectionFileName)"
            }
            else {
                $fullexportdirpath = "$($FilePath)"
            }

            $fullexportpath = ""
            #$fullexportpath = "$($fullfilepathwithoutextension).docx"
            $subpagelinks = ""
            $subpagelinkcountlevel1 = 0
            $subpagelinkcountlevel2 = 0
            $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
             
            if ($global:activateMOCForObsidian)
            {
                # in case multiple pages with the same name exist in a section, postfix the filename. Run after pages 
                if ([System.IO.File]::Exists("$($fullfilepathwithoutextension).md")) {
                    #continue
                    $pagename = "$($pagename)-$global:pageNameRecurrenceCount"
                    $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
                    $global:pageNameRecurrenceCount++ 
                }
            }

            # process for subpage prefixes
            if ($pagelevel -eq 1) {
                $pageprefix = ""
                $previouspagenamelevel1 = $pagename
                $previouspagenamelevel2 = ""
                $previouspagelevel = 1
                $subpagelinkcountlevel1 = 0
                "#### " + $page.name

                if ($global:activateMOCForObsidian -eq 1)
                {
                    $sectionValue = $sectionValue + "`n- [[$($pagename)]]"
                }
            }
            elseif ($pagelevel -eq 2) {
                    $pageprefix = "$($previouspagenamelevel1)"
                    $previouspagenamelevel2 = $pagename
                    $previouspagelevel = 2
                    $subpagelinkcountlevel2 = 0
                    "##### " + $page.name

                    if ($global:activateMOCForObsidian -eq 1)
                    {
                        # Set links to subpages
                        $orig = Get-Content -path "$($fullexportdirpath)\$($previouspagenamelevel1).md" -encoding utf8
                        $subpagelinks = $subpagelinks + "- [[$($pagename)]]"
                        $init = 3+$subpagelinkcountlevel1
                        $ende = $init+1
                        Set-Content -Path "$($fullexportdirpath)\$($previouspagenamelevel1).md" -Value $orig[0..$init], $subpagelinks, $orig[$ende..$orig.Length] -encoding utf8
                        
                        $subpagelinkcountlevel1++
                    }
            }
            elseif ($pagelevel -eq 3) {
                    if ($previouspagelevel -eq 2){
                        $pageprefix = "$($previouspagenamelevel1)$($prefixjoiner)$($previouspagenamelevel2)"
                    }
                    # level 3 under level 1, without a level 2
                    elseif ($previouspagelevel -eq 1) {
                        $pageprefix = "$($previouspagenamelevel1)$($prefixjoiner)"
                    }
                    #and if previous was 3, do nothing/keep previous label
                    $previouspagelevel = 3
                    "####### " + $page.name

                    if ($global:activateMOCForObsidian -eq 1)
                    {
                        # Set links to subpages
                        $orig = Get-Content -path "$($fullexportdirpath)\$($previouspagenamelevel2).md" -encoding utf8
                        $subpagelinks = $subpagelinks + "- [[$($pagename)]]"
                        $init = 3+$subpagelinkcountlevel2
                        $ende = $init+1
                        Set-Content -Path "$($fullexportdirpath)\$($previouspagenamelevel2).md" -Value $orig[0..$init], $subpagelinks, $orig[$ende..$orig.Length] -encoding utf8
                    
                        $subpagelinkcountlevel2++
                    }
            }

            if ($global:activateMOCForObsidian -ne 1)
            {
                #if level 2 or 3 (i.e. has a non-blank pageprefix)
                if ($pageprefix) {
                    #create filename prefixes and filepath if prefixes selected
                    if ($prefixFolders -eq 2) {
                        $pagename = "$($pageprefix)_$($pagename)"
                        $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
                    }
                    #all else/default, create subfolders and filepath if subfolders selected
                    elseif ($global:activateSubDir -eq 1) {
                        New-Item -Path "$($fullexportdirpath)\$($pageprefix)" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
                        $fullexportdirpath = "$($fullexportdirpath)\$($pageprefix)"
                        $fullfilepathwithoutextension = "$($fullexportdirpath)\$($pagename)"
                        $levelsprefix = "../"*($levelsfromroot+$pagelevel-1)+".."
                    }
                }
                elseif ($global:activateSubDir -eq 1)
                {
                    $levelsprefix = "../"*($levelsfromroot)+".."
                }
            }

            saveHyperlinkToObject $page $pagename

            # set media location (central media folder at notebook-level or adjacent to .md file) based on initial user prompt
            if ($medialocation -eq 2) {
                $mediaPath = $fullexportdirpath
                $levelsprefix = ""
            }
            else {
                $mediaPath = $NotebookFilePath
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
            else{
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
                pandoc.exe -f  docx -t $converter-simple_tables-multiline_tables-grid_tables+pipe_tables -i $fullexportpath -o "$($fullfilepathwithoutextension).md" --wrap=none --atx-headers --extract-media="$($mediaPath)"
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
                New-Item -Path "$($mediaPath)" -Name "media" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
                $destfilename = ""
                try {
                    $destfilename = ""

                    if($global:activateGlobalFileName -eq 1) #Uses global file name to avoid naming problems
                    {
                        $global:fileCount++
                        $destfilename = "$($global:filePrefix)$($global:fileCount)$($pageinsertedfile.Extension)"
                    }
                    else {
                        $destfilename = $pageinsertedfile.InsertedFile.preferredName | Remove-InvalidFileNameCharsInsertedFiles    
                    }
                    
                    Copy-Item -Path "$($pageinsertedfile.InsertedFile.pathCache)" -Destination "$($mediaPath)\media\$($destfilename)" -Force
                }
                catch {
                    Write-Host "Error while copying file object '$($pageinsertedfile.InsertedFile.preferredName)' for page '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while copying file object '$($pageinsertedfile.InsertedFile.preferredName)' for page '$($page.name)': $($Error[0].ToString())`r`n"
                }
                # Change MD file Object Name References
                try {
                    $pageinsertedfile2 = $pageinsertedfile.InsertedFile.preferredName.Replace("$","\$").Replace("^","\^").Replace("'","\'")                                
                    ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw -encoding utf8).Replace("$($pageinsertedfile2)", "[$($destfilename)]($($mediaPath)/media/$($destfilename))")) | Set-Content -Path "$($fullfilepathwithoutextension).md" -encoding utf8

                }
                catch {
                    Write-Host "Error while renaming file object name references to '$($pageinsertedfile.InsertedFile.preferredName)' for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while renaming file object name references to '$($pageinsertedfile.InsertedFile.preferredName)' for file '$($page.name)': $($Error[0].ToString())`r`n"
                }
            }
            
            # add YAML
            $orig = Get-Content -path "$($fullfilepathwithoutextension).md" -encoding utf8
            $orig[0] = "# $($page.name)"
            $insert1 = "$($page.dateTime)"
            $insert1 =[Datetime]::ParseExact($insert1, 'yyyy-MM-ddTHH:mm:ss.fffZ', $null)
            $insert1 = $insert1.ToString("yyyy-MM-dd HH:mm:ss
            ")
            $insert2 = "---" 

            if($global:activateMOCForObsidian -eq 1)
            {
                Set-Content -Path "$($fullfilepathwithoutextension).md" -Value $orig[0..0], $insert1, $insert2, $insert2, $orig[6..$orig.Length] -encoding utf8
            }
            else {
                Set-Content -Path "$($fullfilepathwithoutextension).md" -Value $orig[0..0], $insert1, $insert2, $orig[6..$orig.Length] -encoding utf8    
            }
            
            
            #Clear double spaces from bullets and nonbreaking spaces from blank lines
            if ($keepspaces -eq 2 ) {
                #do nothing
            }
            else {
                try {
                    ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw -encoding utf8).Replace(">","").Replace("<","").Replace([char]0x00A0,[char]0x000A).Replace([char]0x000A,[char]0x000A).Replace("`r`n`r`n", "`r`n")) | Set-Content -Path "$($fullfilepathwithoutextension).md" -encoding utf8                  }
                catch {
                    Write-Host "Error while clearing double spaces from file '$($fullfilepathwithoutextension)' : $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while clearing double spaces from file '$($fullfilepathwithoutextension)' : $($Error[0].ToString())`r`n"
                }    
            }
            
            # rename images to have unique names - NoteName-Image#-HHmmssff.xyz
            $timeStamp = (Get-Date -Format HHmmssff).ToString()
            $timeStamp = $timeStamp.replace(':', '')
            $images = Get-ChildItem -Path "$($mediaPath)/media" -Include "*.png", "*.gif", "*.jpg", "*.jpeg" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name.SubString(0,5) -match "image" }
            foreach ($image in $images) {
                $newimageName = ""

                if($global:activateGlobalFileName -eq 1) #Uses global file name to avoid naming problems
                {
                    $global:fileCount++
                    $newimageName = "$($global:filePrefix)$($global:fileCount)$($image.Extension)"
                }
                else {
                    $newimageName = "$($pagename.SubString(0,[math]::min(30,$pagename.length)))-$($image.BaseName)-$($timeStamp)$($image.Extension)"
                }
                
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
                    ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw -encoding utf8).Replace("$($image.Name)", "$($newimageName)")) | Set-Content -Path "$($fullfilepathwithoutextension).md" -encoding utf8
                }
                catch {
                    Write-Host "Error while renaming image file name references to '$($image.Name)' for file '$($page.name)': $($Error[0].ToString())" -ForegroundColor Red
                    $totalerr += "Error while renaming image file name references to '$($image.Name)' for file '$($page.name)': $($Error[0].ToString())`r`n"
                }
            }

            # change MD file Image Path References
            try {
                # Change MD file Image Path References in Markdown
                ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw -encoding utf8).Replace("$($mediaPath.Replace("\","\\"))", "$($levelsprefix)")) | Set-Content -Path "$($fullfilepathwithoutextension).md" -encoding utf8
                # Change MD file Image Path References in HTML
                ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw -encoding utf8).Replace("$($mediaPath)", "$($levelsprefix)")) | Set-Content -Path "$($fullfilepathwithoutextension).md" -encoding utf8
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
                ((Get-Content -path "$($fullfilepathwithoutextension).md" -Raw -encoding utf8).Replace("\",'')) | Set-Content -Path "$($fullfilepathwithoutextension).md" -encoding utf8
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

        if($global:activateMOCForObsidian -eq 1)
        {
            $sectionValue = $sectionValue + "`n---"
            New-Item -Path "$($FilePath)" -Name "$($sectionFileName).md" -ItemType "file" -Value "$($sectionValue)" -ErrorAction SilentlyContinue
            $sectionGroupValue = $sectionGroupValue + "`n- [[$($sectionFileName)]]"
            saveHyperlinkToObject $section $sectionFileName
        }  

    }

    if($global:activateMOCForObsidian -eq 1)
    {
        $sectionGroupValue = $sectionGroupValue + "`n---"
        
        return $sectionGroupValue
    }
}

Function FindNameById($objId, $idEx)
{
    $i = 0
    foreach ($element in $PageValueArray)
    {
        $thisLinkObjId = [regex]::Match($element, $idEx).Groups[1].Value
        if($objId)
        {
            if ($objId -eq $thisLinkObjId)
            {
                return $PageIdArray[$i]
            }
        }
        $i++
    }
    return ""
}

Function MatchLinkToFile ($notesdestpath, $link)
{
    $files = Get-ChildItem -Path $notesdestpath -Recurse -Include "$($link).md"

    if (!$files)
    {
        $files = Get-ChildItem -Path $notesdestpath -Recurse -Include "*$($link)*.md"
    }

    if (!$files)
    {
        #Search without special chars
        $searchTerm = $link -Replace "[^A-Za-z0-9]", '*'
        $files = Get-ChildItem -Path $notesdestpath -Recurse -Include "*$($searchTerm)*.md"
    }

    if($files.Length -gt 0)
    {
        foreach ($file in $files) {
            $link = $file.Name -replace ".md", ''
            break
        }

        if(($files -is [array]) -and ($files.Length -gt 1))
        {
            $link = $link + "???"
        }
    }
    else {
        $link = ""
    }
    return $link
}

Function parseLinkForPattern($file, $linkExp, $nameExp)
{
    # Find links based on regex
    $secLinks = (Get-Content -path $file.FullName -Raw -encoding utf8) | Select-String -Pattern "$($linkExp)" -AllMatches | %{$_.matches} | %{$_.Value}
            
    foreach($link in $secLinks)
    {
        #$url = [regex]::Match($link,"(?=onenote:)(.*?)(?=\))").Groups[1].Value
        $LinkPageId = [regex]::Match($link,"(?<=page-id={)(.*?)(?=})").Groups[1].Value
        #$LinkSectionId = [regex]::Match($link,"(?<=section-id={)(.*?)(?=})").Groups[1].Value
        #$linkObjectId = [regex]::Match($link,"(?<=object-id={)(.*?)(?=})").Groups[1].Value
        #$LinkBasePath = [regex]::Match($link,"(?<=base-path=)(.*?)(?=\))").Groups[1].Value
        
        $linkName = ""

        #Find usig object ids
        if ($LinkPageId)
        {
            $linkName = FindNameById $LinkPageId "(?<=page-id={)(.*?)(?=})"
        }

        if($linkName)
        {
            $linkName = $linkName | Remove-InvalidFileNameChars
        }

        #Find using file search
        if(!$linkName -and $linkExp -and $nameExp)
        {
            $linkName = [regex]::Match($link,$nameExp).Groups[1].Value | Remove-InvalidFileNameChars
            $linkName = $linkName.Replace("%20", " ")
            $linkName = MatchLinkToFile $notesdestpath $linkName
        }

        if ($linkName)
        {
            $append = ""
            $global:parsedLinksMatched++

            if ($linkName.Contains("???"))
            {
                $linkName = $linkName.Replace("???", '')
                $append = "#LinkAmbiguous"
                $global:parsedLinksAmbiguousMatched++
            }
            else {
                $global:parsedLinksMatched++
            }

            ((Get-Content -path $file.FullName -Raw -encoding utf8).Replace("$($link)", "[[$($linkName)]] $($append)")) | Set-Content -Path $file.FullName -encoding utf8
        }
        elseif($nameExp){
            ((Get-Content -path $file.FullName -Raw -encoding utf8).Replace("$($link)", "$($link) #LinkNotResolved")) | Set-Content -Path $file.FullName -Encoding UTF8
            $global:parsedLinksNotMatched++
        }
    }
}

Function ParseLinks($notesdestpath)
{
    $files = Get-ChildItem -Path $notesdestpath -Recurse -Include "*.md"

    foreach ($file in $files) {
        try {
            # All onenote Links
            $linkExp = "(?=\[.*\]\(onenote:)(.*?)(?<=\.one\))"
            $nameExp = ""
            parseLinkForPattern $file $linkExp $nameExp

            # Section Links
            $linkExp = "(?=\[.*\(onenote:#section)(.*?)(?<=\.one\))"
            $nameExp = "(?<=\[)(.*?)(?=\])"
            parseLinkForPattern $file $linkExp $nameExp

            # Page and Paragraph links to page
            $linkExp = "(?=\[.*\(onenote:.*&section)(.*?)(?<=\.one\))"
            $nameExp = "(?<=\]\(onenote:#)(.*?)(?=&section-id)"
            parseLinkForPattern $file $linkExp $nameExp
            
            # Unsesolved onenote links -> [[NAME]]
            $linkExp = "(?=\[)(.*?)(\]\(onenote:)(.*?)(?<=\))"
            $nameExp = "(?<=\[)(.*?)(?=\])"
            parseLinkForPattern $file $linkExp $nameExp
        }
        catch
        {
            Write-Host "Error resolving links for '$($file.name)': $($Error[0].ToString())" -ForegroundColor Red
            $totalerr += "Error resolving links for '$($file.name)': $($Error[0].ToString())`r`n"
        }
    }
    Write-Host "Total links found`t:`t$($global:parsedLinksMatched + $global:parsedLinksNotMatched + $global:parsedLinksAmbiguousMatched)"
    Write-Host "`tMatched`t`t:`t$($global:parsedLinksMatched)"
    Write-Host "`tAmbiguous`t:`t$($global:parsedLinksAmbiguousMatched)"
    Write-Host "`tNot matched`t:`t$($global:parsedLinksNotMatched)"
}

""
"-----------------------------------------------"
# ask for the Notes root path
"Enter the (preferably empty!) folder path (without trailing backslash!) that will contain your resulting Notes structure. ex. 'c:\temp\notes'"
$notesdestpath = Read-Host -Prompt "Entry"
""
"-----------------------------------------------"

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
# prompt MOC for Obsidian
"1: Use Map of Contents Structure - Default"
"`t--For linking in Obsidian--"
"`t`t- All files in the same folder"
"`t`t- Pages for sections and section groups containing links to child pages/sections"
"`t`t- Pages link to subpages (3 levels)"
"2: Select other structure"
[Int]$prefixFolders = 1
[int]$medialocation = 1
$optMOCForObsidian = Read-Host -Prompt "Entry"
if ($optMOCForObsidian -ne 2) 
{
    $global:activateMOCForObsidian = 1
    $global:activateSubDir = 0
    $global:activateGlobalFileName = 1
    $prefixFolders = 2
}
else {
    $global:activateSubDir = 1
    $global:activateGlobalFileName = 0
    $global:activateMOCForObsidian = 0
}

if (!$global:activateMOCForObsidian)
{
    ""
    "-----------------------------------------------"
    # prompt for prefix vs subfolders
    "1: Create folders for subpages (e.g. Page\Subpage.md)- Default"
    "2: Add prefixes for subpages (e.g. Page_Subpage.md)"

    $prefixFolders = Read-Host -Prompt "Entry"
    if ($prefixFolders -eq 2) {
        $prefixFolders = 2 
        $prefixjoiner = "_"
    }
    else {
        $prefixFolders = 1
        $prefixjoiner = "\"
        $global:activateSubDir = 1
    }
}

if (!$global:activateMOCForObsidian)
{
    #prompt for media in single or multiple folders
    ""
    "-----------------------------------------------"
    "1: Images stored in single 'media' folder at Notebook-level (Default)"
    "2: Separate 'media' folder for each folder in the hierarchy"

        $medialocation = Read-Host -Prompt "Entry"
}

#prompt for generic file naming
""
"-----------------------------------------------"
"1: Use generic file names to avoid name errors (File-<Number>) - Default"
"2: Maintain original file names"
if((Read-Host -Prompt "Entry") -eq 2)
{
    $global:activateGlobalFileName = 0
}

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
if ($conversion -eq 2){ $converter = "commonmark"}
elseif ($conversion -eq 3){ $converter = "gfm"}
elseif ($conversion -eq 4){ $converter = "markdown_mmd"}
elseif ($conversion -eq 5){ $converter = "markdown_phpextra"}
elseif ($conversion -eq 6){ $converter = "markdown_strict"}
else { $converter = "markdown"}

#prompt to clear double spaces between bullets
"-----------------------------------------------"
"1: Clear double spaces in bullets - Default"
"2: Keep double spaces"
[int] $keepspaces = Read-Host -Prompt "Entry"

# prompt to clear escape symbols from md files 
"-----------------------------------------------"
"1: Clear '\' symbol escape character from files"
"2: Keep '\' symbol escape"
[int] $keepescape = Read-Host -Prompt "Entry"

# prompt to resolve links for obsidian 
"-----------------------------------------------"
"1: Convert OneNote links to Markdown [[Link]]"
"2: Keep onenote links"
[int] $resolveLinks = Read-Host -Prompt "Entry"


if (Test-Path -Path $notesdestpath) {
    # open OneNote hierarchy
    $OneNote = New-Object -ComObject OneNote.Application
    [xml]$Hierarchy = ""
    $totalerr = ""
    $OneNote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$Hierarchy)

    foreach ($notebook in $Hierarchy.Notebooks.Notebook) {
        " "
        $notebook.Name
        $notebookFileName = "$($notebook.Name)" | Remove-InvalidFileNameChars
        New-Item -Path "$($notesdestpath)\" -Name "$($notebookFileName)" -ItemType "directory" -ErrorAction SilentlyContinue
        $NotebookFilePath = "$($notesdestpath)\$($notebookFileName)"
        $levelsfromroot = 0
       
        New-Item -Path "$($NotebookFilePath)" -Name "docx" -ItemType "directory" -ErrorAction SilentlyContinue

        $notebookContent = ""

        "=============="
        #process any sections that are not in a section group
        if ($global:activateMOCForObsidian -eq 1)
        {
            $notebookContent = (ProcessSections $notebook $NotebookFilePath)[-1]
        }
        else {
            ProcessSections $notebook $NotebookFilePath
        }
        
        #start looping through any top-level section groups in the notebook
        foreach ($sectiongroup1 in $notebook.SectionGroup) {
            $levelsfromroot = 1
            if ($sectiongroup1.isRecycleBin -ne 'true') {
                "# " + $sectiongroup1.Name
                $sectiongroupFileName1 = "$($sectiongroup1.Name)" | Remove-InvalidFileNameChars
                $sectiongroupFilePath1 = ""
                
                if ($global:activateSubDir -eq 1)
                {
                    New-Item -Path "$($notesdestpath)\$($notebookFileName)" -Name "$($sectiongroupFileName1)" -ItemType "directory" -ErrorAction SilentlyContinue 
                    $sectiongroupFilePath1 =  "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)"
                }
                else {
                    $sectiongroupFilePath1 =  "$($notesdestpath)\$($notebookFileName)"
                }
                
                [string]$sectionGroup1Value = ""

                if ($global:activateMOCForObsidian -eq 1)
                {
                    $sectionGroup1Value = (ProcessSections $sectiongroup1 $sectiongroupFilePath1)[-1]
                }
                else {
                    ProcessSections $sectiongroup1 $sectiongroupFilePath1
                }
                
                #start looping through any 2nd level section groups within the 1st level section group
                foreach ($sectiongroup2 in $sectiongroup1.SectionGroup) {
                    $levelsfromroot = 2
                    if ($sectiongroup2.isRecycleBin -ne 'true') {
                        "## " + $sectiongroup2.Name

                        $sectiongroupFileName2 = "$($sectiongroup2.Name)" | Remove-InvalidFileNameChars
                        $sectiongroupFilePath2 = ""

                        if ($global:activateSubDir -eq 1)
                        {
                            New-Item -Path "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)" -Name "$($sectiongroupFileName2)" -ItemType "directory" -ErrorAction SilentlyContinue
                            $sectiongroupFilePath2 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)"
                        }
                        else {
                            $sectiongroupFilePath2 = "$($notesdestpath)\$($notebookFileName)"
                        }
                        
                        [string]$sectionGroup2Value = ""

                        if ($global:activateMOCForObsidian -eq 1)
                        {
                            $sectionGroup1Value = $sectionGroup1Value + "`n- [[$($sectiongroupFileName2)]]"
                            $sectionGroup2Value = (ProcessSections $sectiongroup2 $sectiongroupFilePath2)[-1]
                        }
                        else {
                            ProcessSections $sectiongroup2 $sectiongroupFilePath2
                        }
                        
                        #start looping through any 2nd level section groups within the 1st level section group
                        foreach ($sectiongroup3 in $sectiongroup2.SectionGroup) {
                            $levelsfromroot = 3
                            if ($sectiongroup3.isRecycleBin -ne 'true') {
                                "### " + $sectiongroup3.Name
                                
                                $sectiongroupFileName3 = "$($sectiongroup3.Name)" | Remove-InvalidFileNameChars
                                $sectiongroupFilePath3 = ""

                                if ($global:activateSubDir -eq 1)
                                {
                                    New-Item -Path "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)" -Name "$($sectiongroupFileName3)" -ItemType "directory" -ErrorAction SilentlyContinue
                                    $sectiongroupFilePath3 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)"
                                }
                                else {
                                    $sectiongroupFilePath3 = "$($notesdestpath)\$($notebookFileName)"   
                                }

                                [string]$sectionGroup3Value = ""
                                
                                if ($global:activateMOCForObsidian -eq 1)
                                {
                                    $sectionGroup2Value = $sectionGroup2Value + "`n- [[$($sectiongroupFileName3)]]"
                                    $sectionGroup3Value =  (ProcessSections $sectiongroup3 $sectiongroupFilePath3)[-1]
                                }
                                ProcessSections $sectiongroup3 $sectiongroupFilePath3
                                
                                #start looping through any 2nd level section groups within the 1st level section group
                                foreach ($sectiongroup4 in $sectiongroup3.SectionGroup) {
                                    $levelsfromroot = 4
                                    if ($sectiongroup4.isRecycleBin -ne 'true') {
                                        "#### " + $sectiongroup4.Name
                                        $sectiongroupFileName4 = "$($sectiongroup4.Name)" | Remove-InvalidFileNameChars
                                        
                                        $sectiongroupFilePath4 = ""
                                        if ($global:activateSubDir -eq 1)
                                        {
                                            New-Item -Path "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)" -Name "$($sectiongroupFileName4)" -ItemType "directory" -ErrorAction SilentlyContinue
                                            $sectiongroupFilePath4 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)\$($sectiongroupFileName4)"
                                        }
                                        else {
                                            $sectiongroupFilePath4 = "$($notesdestpath)\$($notebookFileName)"   
                                        }

                                        [string]$sectionGroup4Value = ""
                                
                                        if ($global:activateMOCForObsidian -eq 1)
                                        {
                                            $sectionGroup3Value = $sectionGroup3Value + "`n- [[$($sectiongroupFileName4)]]"
                                            $sectionGroup4Value =  (ProcessSections $sectiongroup4 $sectiongroupFilePath4)[-1]
                                        }
                                        else {
                                            ProcessSections $sectiongroup4 $sectiongroupFilePath4
                                        }
                                                                                
                                        #start looping through any 2nd level section groups within the 1st level section group
                                        foreach ($sectiongroup5 in $sectiongroup4.SectionGroup) {
                                            $levelsfromroot = 5
                                            if ($sectiongroup5.isRecycleBin -ne 'true') {
                                                "#### " + $sectiongroup5.Name
                                                $sectiongroupFileName5 = "$($sectiongroup5.Name)" | Remove-InvalidFileNameChars
                                                $sectiongroupFilePath5 = ""
                                                if ($global:activateSubDir -eq 1)
                                                {
                                                    New-Item -Path "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)\$($sectiongroupFileName4)" -Name "$($sectiongroupFileName5)" -ItemType "directory" -ErrorAction SilentlyContinue
                                                    $sectiongroupFilePath5 = "$($notesdestpath)\$($notebookFileName)\$($sectiongroupFileName1)\$($sectiongroupFileName2)\$($sectiongroupFileName3)\$($sectiongroupFileName4)\\$($sectiongroupFileName5)"
                                                }
                                                else {
                                                    $sectiongroupFilePath5 = "$($notesdestpath)\$($notebookFileName)"   
                                                }

                                                [string]$sectionGroup5Value = ""
                                        
                                                if ($global:activateMOCForObsidian -eq 1)
                                                {
                                                    $sectionGroup4Value = $sectionGroup4Value + "`n- [[$($sectiongroupFileName5)]]"
                                                    $sectionGroup5Value =  (ProcessSections $sectiongroup5 $sectiongroupFilePath5)[-1]
                                                }
                                                else {
                                                    ProcessSections $sectiongroup5 $sectiongroupFilePath5
                                                }
                                            }
                                        }
                                        
                                        if ($global:activateMOCForObsidian -eq 1)
                                        {
                                            New-Item -Path "$($notesdestpath)" -Name "$($sectiongroupFileName4).md" -ItemType "file" -Value "$($sectionGroup4Value)" -ErrorAction SilentlyContinue  
                                        }
                                    }
                                }
                                if ($global:activateMOCForObsidian -eq 1)
                                {
                                    New-Item -Path "$($notesdestpath)" -Name "$($sectiongroupFileName3).md" -ItemType "file" -Value "$($sectionGroup3Value)" -ErrorAction SilentlyContinue
                                }
                            }
                        }
                        if ($global:activateMOCForObsidian -eq 1)
                        {
                            New-Item -Path "$($notesdestpath)" -Name "$($sectiongroupFileName2).md" -ItemType "file" -Value "$($sectionGroup2Value)" -ErrorAction SilentlyContinue
                        }
                    }
                }
                if ($global:activateMOCForObsidian -eq 1)
                {
                    New-Item -Path "$($sectiongroupFilePath1)" -Name "$($sectiongroupFileName1).md" -ItemType "file" -Value "$($sectionGroup1Value)" -ErrorAction SilentlyContinue
                    $notebookContent = $notebookContent + "`n- [[$($sectiongroupFileName1)]]"
                }
            }
        } 
        
        if ($global:activateMOCForObsidian -eq 1)
        {
            $notebookContent = $notebookContent + "`n---"
            New-Item -Path "$($NotebookFilePath)" -Name "$($notebookFileName).md" -ItemType "file" -Value "$($notebookContent)" -ErrorAction SilentlyContinue
        }
    }

    #Parse links
    if($resolveLinks -ne 2)
    {
        parseLinks $notesdestpath
    }

    # release OneNote hierarchy
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OneNote)
    Remove-Variable OneNote
    $totalerr
}
else {
Write-Host "This path is NOT valid"
}
