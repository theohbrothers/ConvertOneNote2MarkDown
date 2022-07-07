$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut" -Exit

Set-StrictMode -version latest

Describe "Validate-Dependencies" -Tag 'Unit' {

    Context 'Behavior' {

        It "Warn when GAC assemblies are missing (Windows)" {
            Mock Get-Item {}
            Mock Write-Warning { 'a warning' }
            Mock Get-Command { $true }

            if ($env:OS -imatch 'Windows') {
                $warning = Validate-Dependencies
                $warning | Should -Match 'a warning'
            }else {
                $true
            }
        }

        It "Throws exception when pandoc is not found" {
            Mock Get-Command {}

            { Validate-Dependencies } | Should -Throw
        }

    }

}

Describe "Compile-Configuration" -Tag 'Unit' {

    Context 'Behavior' {

        It "Compiles configuration from config file: cast input as expected type, normalize path, trim string, and fallback on default values on empty input" {
            Mock Test-Path { $true }
            Mock Get-Content {
                # Fake content of a config.ps1
                @'
$dryrun = 1
$notesdestpath = 'c:\temp\notes\/ ' # Deliberately add a trailing slah(es) and space
$targetNotebook = '   ' # Deliberately add extra spaces
$usedocx = '1'
# $keepdocx = 1 # Deliberately omit a configuration option from
$prefixFolders = 1
$medialocation = 1
$conversion = 'markdown-simple_tables-multiline_tables-grid_tables+pipe_tables' # Default
$headerTimestampEnabled = 1
$keepspaces = 1
$keepescape = 1
$newlineCharacter = 1
'@
            }
            $expectedConfig = Get-DefaultConfiguration

            $config = Compile-Configuration

            foreach ($k in $config.Keys) {
                $config[$k]['value'] | Should -Be  $expectedConfig[$k]['value']
            }
        }

        It "Compiles configuration from interactive prompts: cast input as expected type, and fallback on default values on empty input" {
            $fakeUserInputs = Get-DefaultConfiguration
            $fakeUserInputs['notesdestpath']['value'] = 'c:\foo\bar' # Should be intact
            $fakeUserInputs['targetNotebook']['value'] = $null # Should fallback on default value
            $fakeUserInputs['usedocx']['value'] = '1' # Should be casted to an int
            $fakeUserInputs['keepdocx']['value'] = $null # Should fallback on default value
            Mock Test-Path { $false }
            Mock Read-Host {
                $typeName = [Microsoft.PowerShell.ToStringCodeMethods]::Type($config[$key]['default'].GetType())
                Invoke-Expression ('$fakeUserInputs[$key]["value"] -as ' + "[$typeName]")
            }

            $expectedConfig = Get-DefaultConfiguration
            $expectedConfig['notesdestpath']['value'] = 'c:\foo\bar'

            $config = Compile-Configuration 6>$null

            foreach ($k in $config.Keys) {
                $config[$k]['value'] | Should -Be $expectedConfig[$k]['value']
            }
        }

    }

}

Describe "Validate-Configuration" -Tag 'Unit' {

    Context 'Parameters' {

        It "Throw exception when no parameters are specified" {
            { Validate-Configuration } | Should -Throw 'No input parameters specified.'
        }

        It "Accept pipeline input" {
            $config = @{}
            Mock Get-DefaultConfiguration { @{} }

            { $config | Validate-Configuration } | Should -Not -Throw
        }

    }

    Context 'Behavior' {

        It "Throws on missing config option" {
            $config = Get-DefaultConfiguration
            $config['notesdestpath'] = $null

            { $config | Validate-Configuration } | Should -Throw 'Missing or invalid configuration option'

            $config = Get-DefaultConfiguration
            $config['notesdestpath']['value'] = $null

            { $config | Validate-Configuration } | Should -Throw 'Missing or invalid configuration option'
        }

        It "Throws on invalid config option type" {
            $config = Get-DefaultConfiguration
            $config['notesdestpath']['value']= 12345

            { $config | Validate-Configuration } | Should -Throw 'Expected a value of type'
        }

        It "Throws on config option being a path that does not exist" {
            $config = Get-DefaultConfiguration
            Mock Test-Path { $false }

            { $config | Validate-Configuration } | Should -Throw 'does not exist, or is a file'
        }

        It "Throws on config option falling outside of valid range of integer values" {
            # Valid
            1..2 | % {
                $config = Get-DefaultConfiguration
                Mock Test-Path { $true }
                $config['usedocx']['value'] = $_

                { $config | Validate-Configuration } | Should -Not -Throw
            }

            # Invalid
            0,3 | % {
                $config = Get-DefaultConfiguration
                Mock Test-Path { $true }
                $config['usedocx']['value'] = $_

                { $config | Validate-Configuration } | Should -Throw 'The value must be between'
            }
        }

    }

}

Describe "Print-Configuration" -Tag 'Unit' {

    Context 'Parameters' {

        It "Throw exception when no parameters are specified" {
            { Validate-Configuration } | Should -Throw 'No input parameters specified.'
        }

        It "Accept pipeline input" {
            $config = @{}
            Mock Get-DefaultConfiguration { @{} }

            { $config | Validate-Configuration } | Should -Not -Throw
        }

    }

    Context 'Behavior' {

        It "Prints configuration to Write-Host stream" {
            $config = Get-DefaultConfiguration

            $result = $config | Print-Configuration 6>&1
            $result | % {
                $_ | Should -match '[^:]+:[^:]+'
            }
        }

    }

}

Describe "Remove-InvalidFileNameChars" -Tag 'Unit' {

    Context 'Parameters' {

        It "Accept pipeline input" {
            $fileName = 'foo'

            { $fileName | Remove-InvalidFileNameChars } | Should -Not -Throw
        }

    }

    Context 'Behavior' {

        It "Should remove invalid characters" {
            $fileName = 'foo[bar]'
            $expectedFileName = 'foo(bar)'

            $result = Remove-InvalidFileNameChars -Name $fileName

            $result | Should -Be $expectedFileName
        }

        It "Should replace spaces with dashes by default" {
            $fileName = 'foo bar'
            $expectedFileName = 'foo-bar'

            $result = Remove-InvalidFileNameChars -Name $fileName

            $result | Should -Be $expectedFileName
        }

        It "Should not replace spaces with dashes if -KeepPathSpaces" {
            $fileName = 'foo bar'
            $expectedFileName = 'foo bar'

            $result = Remove-InvalidFileNameChars -Name $fileName -KeepPathSpaces

            $result | Should -Be $expectedFileName
        }

        It "Should trim boundary whitespaces" {
            $fileName = ' foobar '
            $expectedFileName = 'foobar'

            $result = Remove-InvalidFileNameChars -Name $fileName -KeepPathSpaces

            $result | Should -Be $expectedFileName
        }

    }

}

Describe "Truncate-PathFileName" -Tag 'Unit' {

    Context 'Parameters' {

        It "Accept pipeline input" {
            $path = 'foo'

            { $path | Truncate-PathFileName } | Should -Not -Throw
        }

        It "Throw if -Length is greater than 255" {
            $drive = "C:$( [io.path]::DirectorySeparatorChar )" # E.g. C:\
            $path = $drive  + ("a" * 100)
            $length = 500

            { $path | Truncate-PathFileName -Length $length } | Should -Throw 'greater than the maximum allowed range of 255'
        }

    }

    Context 'Behavior' {

        It "Should not truncate a given path's filename to 255 characters if given string is within the limit" {
            $drive = "C:$( [io.path]::DirectorySeparatorChar )" # E.g. C:\
            $path = $drive + ("a" * 100) # E.g. C:\aaaaa....
            $expectedPath = $drive  + ("a" * 100) # E.g. C:\aaaaa....

            $path | Truncate-PathFileName | Should -Be $expectedPath
        }

        It "Should truncate a given path's filename to 255 characters by default" {
            $drive = "C:$( [io.path]::DirectorySeparatorChar )" # E.g. C:\
            $path = $drive  + ("a" * 1000)
            $expectedPath = $drive  + ("a" * 255)

            $path | Truncate-PathFileName | Should -Be $expectedPath
        }

        It "Should truncate a given path's filename to a specified -Length" {
            $drive = "C:$( [io.path]::DirectorySeparatorChar )" # E.g. C:\
            $path = $drive  + ("a" * 1000)
            $length = 123
            $expectedPath = $drive  + ("a" * 123)

            $path | Truncate-PathFileName -Length $length | Should -Be $expectedPath
        }

    }

}

Describe "Remove-InvalidFileNameCharsInsertedFiles" -Tag 'Unit' {

    Context 'Parameters' {

        It "Accept pipeline input" {
            $fileName = 'foo'

            { $fileName | Remove-InvalidFileNameCharsInsertedFiles } | Should -Not -Throw
        }

    }

    Context 'Behavior' {

        It "Should remove invalid characters" {
            $fileName = "foo#$%^*[]'<>!@{};bar"
            $expectedFileName = if ($env:OS -match 'windows') { 'foo---bar' } else { 'foobar' } # Compared to unix, Windows does not permit '*', '<', '>'

            $result = Remove-InvalidFileNameCharsInsertedFiles -Name $fileName

            $result | Should -Be $expectedFileName
        }

        It "Should replace spaces with dashes by default" {
            $fileName = 'foo bar'
            $expectedFileName = 'foo-bar'

            $result = Remove-InvalidFileNameCharsInsertedFiles -Name $fileName

            $result | Should -Be $expectedFileName
        }

        It "Should not replace spaces with dashes if -KeepPathSpaces" {
            $fileName = 'foo bar'
            $expectedFileName = 'foo bar'

            $result = Remove-InvalidFileNameCharsInsertedFiles -Name $fileName -KeepPathSpaces

            $result | Should -Be $expectedFileName
        }

        It "Should trim boundary whitespaces" {
            $fileName = ' foobar '
            $expectedFileName = 'foobar'

            $result = Remove-InvalidFileNameCharsInsertedFiles -Name $fileName -KeepPathSpaces

            $result | Should -Be $expectedFileName
        }

    }

}


Describe "Encode-Markdown" -Tag 'Unit' {

    Context 'Parameters' {

        It "Accept pipeline input" {
            $content = 'foo'

            { $content | Encode-Markdown } | Should -Not -Throw
        }

    }

    Context 'Behavior' {

        $content = '\*_{}[]()#+-.!`'
        It "Should markdown-encode given content" {
            $expectedContent = '\\\*\_\{\}\[\]\(\)\#\+\-\.\!```'

            $result = Encode-Markdown -Name $content

            $result | Should -Be $expectedContent
        }

        It "Should markdown-encode given content (URIs)" {
            $expectedContent = '\*_{}\[\]\(\)#+-.!`'

            $result = Encode-Markdown -Name $content -Uri

            $result | Should -Be $expectedContent
        }

    }

}

Describe "New-OneNoteConnection" -Tag 'Unit' {

    Context 'Behavior' {

        # It "Should throw on failed a onenote conection" {
        #     Mock New-Object { $false }
        #     Mock Add-Type { $false }

        #     { New-OneNoteConnection } | Should -Throw 'Failed to make connection to OneNote.'
        # }

        It "Should make a onenote conection" {
            if ($PSVersionTable.PSVersion.Major -le 5) {
                Mock New-Object { 'onenote' }

                $result = New-OneNoteConnection

                $result | Should -Be 'onenote'
            }else {
                # Mock Add-Type { $true }

                # $result = New-OneNoteConnection

                $true
            }
        }

    }

}

function Get-FakeOneNoteHierarchyWithEmptySectionGroupsAndSectionsAndPages {
    # Sample outerXML of a hierarchy object.
    # 1) Section is empty, in the notebook base
    # 2) Section Group is empty, in the notebook base
    $hierarchy = @'
<?xml version="1.0"?>
<one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
    <one:Notebook name="test" nickname="test" ID="{38E47DAB-211E-4EC1-85F1-129656A9D2CE}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/" lastModifiedTime="2021-08-06T16:27:58.000Z" color="#ADE792">
        <one:Section name="s0" ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/s0.one" lastModifiedTime="2021-08-06T16:08:25.000Z" color="#8AA8E4">
        </one:Section>
        <one:SectionGroup name="OneNote_RecycleBin" ID="{1298D961-43A6-46E4-81FC-B4FD9F87755C}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/OneNote_RecycleBin/" lastModifiedTime="2021-08-06T16:27:58.000Z" isRecycleBin="true">
        </one:SectionGroup>
        <one:SectionGroup name="g0" ID="{9570CCF6-17C2-4DCE-83A0-F58AE8914E29}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/g9/" lastModifiedTime="2021-08-06T15:49:20.000Z">
        </one:SectionGroup>
    </one:Notebook>
</one:Notebooks>
'@ -as [xml]
    $hierarchy
}

function Get-FakeOneNoteHierarchy {
    # Sample outerXML of a hierarchy object. Here we have two identical notebooks: 'test' and 'test2' with a simple nested structure
    # 1) The 1st to 8th are in the notebook base. The 6th - 8th pages test a 3rd level page preceded by a 1st level page.
    # 2) A copy of 1), but nested 1 level
    # 3) A copy of 1), but nested 2 levels
    $hierarchy = @'
<?xml version="1.0"?>
<one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
    <one:Notebook name="test" nickname="test" ID="{38E47DAB-211E-4EC1-85F1-129656A9D2CE}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/" lastModifiedTime="2021-08-06T16:27:58.000Z" color="#ADE792">
        <one:Section name="s0" ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/s0.one" lastModifiedTime="2021-08-06T16:08:25.000Z" color="#8AA8E4">
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096990}" name="p0.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096991}" name="p0.1 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096992}" name="p0.2 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:46:36.000Z" pageLevel="3" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096993}" name="p0.3 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096994}" name="p0.4 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096995}" name="p0.5 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:46:36.000Z" pageLevel="3" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096996}" name="p0.6 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096997}" name="p0.7 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
            </one:Section>
        <one:SectionGroup name="OneNote_RecycleBin" ID="{1298D961-43A6-46E4-81FC-B4FD9F87755C}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/OneNote_RecycleBin/" lastModifiedTime="2021-08-06T16:27:58.000Z" isRecycleBin="true">
            <one:Section name="Deleted Pages" ID="{4E51704F-4F96-4D17-A453-EE11A4BEEBD8}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/OneNote_RecycleBin/OneNote_DeletedPages.one" lastModifiedTime="2021-08-06T15:57:18.000Z" color="#E1E1E1" isInRecycleBin="true" isDeletedPages="true">
                <one:Page ID="{4E51704F-4F96-4D17-A453-EE11A4BEEBD8}{1}{E19471045090380451344020177961181143029428871}" name="Untitled page" dateTime="2021-08-06T00:44:05.000Z" lastModifiedTime="2021-08-06T00:44:05.000Z" pageLevel="1" isInRecycleBin="true" />
            </one:Section>
        </one:SectionGroup>
        <one:SectionGroup name="g0" ID="{9570CCF6-17C2-4DCE-83A0-F58AE8914E29}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/g9/" lastModifiedTime="2021-08-06T15:49:20.000Z">
            <one:Section name="s1" ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/g9/s1.one" lastModifiedTime="2021-08-06T15:49:13.000Z" color="#F5F96F">
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E1954561882093339466822011309233174218559400}" name="p1.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:49:13.000Z" pageLevel="1" />
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E1954561882093339466822011309233174218559401}" name="p1.1 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:44:24.000Z" pageLevel="2" />
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E1954561882093339466822011309233174218559402}" name="p1.2 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:44:29.000Z" pageLevel="3" />
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E1954561882093339466822011309233174218559403}" name="p1.3 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:44:24.000Z" pageLevel="2" />
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E1954561882093339466822011309233174218559404}" name="p1.4 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:49:13.000Z" pageLevel="1" />
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E1954561882093339466822011309233174218559405}" name="p1.5 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:44:29.000Z" pageLevel="3" />
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E1954561882093339466822011309233174218559406}" name="p1.6 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:44:24.000Z" pageLevel="2" />
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E1954561882093339466822011309233174218559407}" name="p1.7 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:49:13.000Z" pageLevel="1" />
                </one:Section>
            <one:SectionGroup name="g1" ID="{9FD73490-F779-4556-BEBF-2185A1938883}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/g9/g1/" lastModifiedTime="2021-08-06T15:49:20.000Z">
                <one:Section name="s2" ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/g9/g1/s2.one" lastModifiedTime="2021-08-06T15:49:20.000Z" color="#ADE792">
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E1946648658300603829922011864041866514264980}" name="p2.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E1946648658300603829922011864041866514264981}" name="p2.1 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:45:37.000Z" pageLevel="2" />
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E1946648658300603829922011864041866514264982}" name="p2.2 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:45:38.000Z" pageLevel="3" />
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E1946648658300603829922011864041866514264983}" name="p2.3 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:45:37.000Z" pageLevel="2" />
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E1946648658300603829922011864041866514264984}" name="p2.4 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E1946648658300603829922011864041866514264985}" name="p2.5 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:45:38.000Z" pageLevel="3" />
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E1946648658300603829922011864041866514264986}" name="p2.6 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:45:37.000Z" pageLevel="2" />
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E1946648658300603829922011864041866514264987}" name="p2.7 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                    </one:Section>
            </one:SectionGroup>
        </one:SectionGroup>
    </one:Notebook>
    <one:Notebook name="test2" nickname="test2" ID="{262450BF-461F-4951-92D3-80C3FC76D27B}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/" lastModifiedTime="2021-08-06T16:28:08.000Z" color="#91BAAE" isCurrentlyViewed="true">
        <one:Section name="s0" ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/s0.one" lastModifiedTime="2021-08-06T16:09:11.000Z" color="#8AA8E4">
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E1955511494331087760101966629430768838010100}" name="p0.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:09:11.000Z" pageLevel="1" />
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E1955511494331087760101966629430768838010101}" name="p0.1 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E1955511494331087760101966629430768838010102}" name="p0.2 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:46:36.000Z" pageLevel="3" />
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E1955511494331087760101966629430768838010103}" name="p0.3 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E1955511494331087760101966629430768838010104}" name="p0.4 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:09:11.000Z" pageLevel="1" />
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E1955511494331087760101966629430768838010105}" name="p0.5 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:46:36.000Z" pageLevel="3" />
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E1955511494331087760101966629430768838010106}" name="p0.6 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E1955511494331087760101966629430768838010107}" name="p0.7 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:09:11.000Z" pageLevel="1" />
            </one:Section>
        <one:SectionGroup name="OneNote_RecycleBin" ID="{1908052E-BCA5-491C-AC6D-6A0D0914B94A}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/OneNote_RecycleBin/" lastModifiedTime="2021-08-06T16:28:08.000Z" isRecycleBin="true" isCurrentlyViewed="true">
            <one:Section name="Deleted Pages" ID="{E136ED4A-ED9F-4B1B-A562-183C91FF483A}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/OneNote_RecycleBin/OneNote_DeletedPages.one" lastModifiedTime="2021-08-06T15:57:46.000Z" color="#E1E1E1" isCurrentlyViewed="true" isInRecycleBin="true" isDeletedPages="true">
                <one:Page ID="{E136ED4A-ED9F-4B1B-A562-183C91FF483A}{1}{E19521355828597401344720101901508363770787001}" name="Untitled page" dateTime="2021-08-06T15:57:35.000Z" lastModifiedTime="2021-08-06T15:57:35.000Z" pageLevel="1" isInRecycleBin="true" isCurrentlyViewed="true" />
            </one:Section>
        </one:SectionGroup>
        <one:SectionGroup name="g0" ID="{12BCC0AF-9077-4795-89B3-92731DA42BF1}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/g0/" lastModifiedTime="2021-08-06T15:45:38.000Z">
            <one:Section name="s1" ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/g0/s1.one" lastModifiedTime="2021-08-06T15:44:31.000Z" color="#F5F96F">
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E1952215876640442054192014463679294726156730}" name="p1.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:43:18.000Z" pageLevel="1" />
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E1952215876640442054192014463679294726156731}" name="p1.1 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:44:24.000Z" pageLevel="2" />
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E1952215876640442054192014463679294726156732}" name="p1.2 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:44:29.000Z" pageLevel="3" />
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E1952215876640442054192014463679294726156733}" name="p1.3 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:44:24.000Z" pageLevel="2" />
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E1952215876640442054192014463679294726156734}" name="p1.4 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:43:18.000Z" pageLevel="1" />
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E1952215876640442054192014463679294726156735}" name="p1.5 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:44:29.000Z" pageLevel="3" />
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E1952215876640442054192014463679294726156736}" name="p1.6 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:44:24.000Z" pageLevel="2" />
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E1952215876640442054192014463679294726156737}" name="p1.7 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:43:18.000Z" pageLevel="1" />
            </one:Section>
            <one:SectionGroup name="g1" ID="{70DCA046-102D-43B3-A39B-EA8A6B7D7CAF}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/g0/g1/" lastModifiedTime="2021-08-06T15:45:38.000Z">
                <one:Section name="s2" ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/g0/g1/s2.one" lastModifiedTime="2021-08-06T15:45:38.000Z" color="#ADE792">
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297670}" name="p2.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297671}" name="p2.1 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:45:37.000Z" pageLevel="2" />
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297672}" name="p2.2 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:45:38.000Z" pageLevel="3" />
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297673}" name="p2.3 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:45:37.000Z" pageLevel="2" />
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297674}" name="p2.4 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297675}" name="p2.5 test" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:45:38.000Z" pageLevel="3" />
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297676}" name="p2.6 test" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:45:37.000Z" pageLevel="2" />
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297677}" name="p2.7 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                </one:Section>
            </one:SectionGroup>
        </one:SectionGroup>
    </one:Notebook>
</one:Notebooks>
'@ -as [xml]
    $hierarchy
}

function Get-FakeOneNoteHierarchyWithDuplicatePageNames {
    # Sample outerXML of a hierarchy object. Here we have two identical notebooks: 'test' and 'test2' with a simple nested structure, each with 9 pages, in groups of 3:
    # 1) The 1st, 2nd, and 3rd pages are identically named, in the notebook base
    # 2) A copy of 1), but nested 1 level
    # 3) A copy of 1), but nested 2 levels
    $hierarchy = @'
<?xml version="1.0"?>
<one:Notebooks xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
    <one:Notebook name="test" nickname="test" ID="{38E47DAB-211E-4EC1-85F1-129656A9D2CE}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/" lastModifiedTime="2021-08-06T16:27:58.000Z" color="#ADE792">
        <one:Section name="s0" ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/s0.one" lastModifiedTime="2021-08-06T16:08:25.000Z" color="#8AA8E4">
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096991}" name="p0.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096992}" name="p0.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096993}" name="p0.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
        </one:Section>
        <one:SectionGroup name="OneNote_RecycleBin" ID="{1298D961-43A6-46E4-81FC-B4FD9F87755C}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/OneNote_RecycleBin/" lastModifiedTime="2021-08-06T16:27:58.000Z" isRecycleBin="true">
            <one:Section name="Deleted Pages" ID="{4E51704F-4F96-4D17-A453-EE11A4BEEBD8}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/OneNote_RecycleBin/OneNote_DeletedPages.one" lastModifiedTime="2021-08-06T15:57:18.000Z" color="#E1E1E1" isInRecycleBin="true" isDeletedPages="true">
                <one:Page ID="{4E51704F-4F96-4D17-A453-EE11A4BEEBD8}{1}{E19471045090380451344020177961181143029428871}" name="Untitled page" dateTime="2021-08-06T00:44:05.000Z" lastModifiedTime="2021-08-06T00:44:05.000Z" pageLevel="1" isInRecycleBin="true" />
            </one:Section>
        </one:SectionGroup>
        <one:SectionGroup name="g0" ID="{9570CCF6-17C2-4DCE-83A0-F58AE8914E29}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/g9/" lastModifiedTime="2021-08-06T15:49:20.000Z">
            <one:Section name="s1" ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/g9/s1.one" lastModifiedTime="2021-08-06T15:49:13.000Z" color="#F5F96F">
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E195456188209333946682201130923317421855940917}" name="p1.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:49:13.000Z" pageLevel="1" />
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E195456188209333946682201130923317421855940918}" name="p1.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:49:13.000Z" pageLevel="1" />
                <one:Page ID="{BE566C4F-73DC-43BD-AE7A-1954F8B22C2A}{1}{E195456188209333946682201130923317421855940919}" name="p1.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:49:13.000Z" pageLevel="1" />
            </one:Section>
            <one:SectionGroup name="g1" ID="{9FD73490-F779-4556-BEBF-2185A1938883}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/g9/g1/" lastModifiedTime="2021-08-06T15:49:20.000Z">
                <one:Section name="s2" ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/g9/g1/s2.one" lastModifiedTime="2021-08-06T15:49:20.000Z" color="#ADE792">
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E194664865830060382992201186404186651426498715}" name="p2.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E194664865830060382992201186404186651426498716}" name="p2.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                    <one:Page ID="{0ED877BC-2887-4FF7-B5D4-394608754507}{1}{E194664865830060382992201186404186651426498717}" name="p2.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                </one:Section>
            </one:SectionGroup>
        </one:SectionGroup>
    </one:Notebook>
    <one:Notebook name="test2" nickname="test2" ID="{262450BF-461F-4951-92D3-80C3FC76D27B}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/" lastModifiedTime="2021-08-06T16:28:08.000Z" color="#91BAAE" isCurrentlyViewed="true">
        <one:Section name="s0" ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/s0.one" lastModifiedTime="2021-08-06T16:09:11.000Z" color="#8AA8E4">
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E19555114943310877601019666294307688380101014}" name="p0.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:09:11.000Z" pageLevel="1" />
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E19555114943310877601019666294307688380101015}" name="p0.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:09:11.000Z" pageLevel="1" />
            <one:Page ID="{F8F6B54D-769A-43DF-9FD0-ED9F51054880}{1}{E19555114943310877601019666294307688380101016}" name="p0.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:09:11.000Z" pageLevel="1" />
        </one:Section>
        <one:SectionGroup name="OneNote_RecycleBin" ID="{1908052E-BCA5-491C-AC6D-6A0D0914B94A}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/OneNote_RecycleBin/" lastModifiedTime="2021-08-06T16:28:08.000Z" isRecycleBin="true" isCurrentlyViewed="true">
            <one:Section name="Deleted Pages" ID="{E136ED4A-ED9F-4B1B-A562-183C91FF483A}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/OneNote_RecycleBin/OneNote_DeletedPages.one" lastModifiedTime="2021-08-06T15:57:46.000Z" color="#E1E1E1" isCurrentlyViewed="true" isInRecycleBin="true" isDeletedPages="true">
                <one:Page ID="{E136ED4A-ED9F-4B1B-A562-183C91FF483A}{1}{E19521355828597401344720101901508363770787001}" name="Untitled page" dateTime="2021-08-06T15:57:35.000Z" lastModifiedTime="2021-08-06T15:57:35.000Z" pageLevel="1" isInRecycleBin="true" isCurrentlyViewed="true" />
            </one:Section>
        </one:SectionGroup>
        <one:SectionGroup name="g0" ID="{12BCC0AF-9077-4795-89B3-92731DA42BF1}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/g0/" lastModifiedTime="2021-08-06T15:45:38.000Z">
            <one:Section name="s1" ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/g0/s1.one" lastModifiedTime="2021-08-06T15:44:31.000Z" color="#F5F96F">
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E19522158766404420541920144636792947261567311}" name="p1.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:43:18.000Z" pageLevel="1" />
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E19522158766404420541920144636792947261567312}" name="p1.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:43:18.000Z" pageLevel="1" />
                <one:Page ID="{66C286AD-A732-4921-BFC3-DB419E018FC5}{1}{E19522158766404420541920144636792947261567313}" name="p1.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:43:18.000Z" pageLevel="1" />
            </one:Section>
            <one:SectionGroup name="g1" ID="{70DCA046-102D-43B3-A39B-EA8A6B7D7CAF}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/g0/g1/" lastModifiedTime="2021-08-06T15:45:38.000Z">
                <one:Section name="s2" ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test2/g0/g1/s2.one" lastModifiedTime="2021-08-06T15:45:38.000Z" color="#ADE792">
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297671}" name="p2.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297672}" name="p2.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                    <one:Page ID="{659C0A68-EEC9-454E-AEB6-34AFE4D8C395}{1}{E1950041523555344809591971866755563186297673}" name="p2.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T15:45:35.000Z" pageLevel="1" />
                </one:Section>
            </one:SectionGroup>
        </one:SectionGroup>
    </one:Notebook>
</one:Notebooks>
'@ -as [xml]
    $hierarchy
}

function Get-FakeOneNotePageContent {
    # Sample outerXML of a page object
    $page = @'
<?xml version="1.0"?>
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote" ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096991}" name="p0.0 test" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:46.000Z" pageLevel="1" selected="partial" lang="en-US">
    <one:QuickStyleDef index="0" name="PageTitle" fontColor="automatic" highlightColor="automatic" font="Calibri Light" fontSize="20.0" spaceBefore="0.0" spaceAfter="0.0" />
    <one:QuickStyleDef index="1" name="p" fontColor="automatic" highlightColor="automatic" font="Calibri" fontSize="11.0" spaceBefore="0.0" spaceAfter="0.0" />
    <one:PageSettings RTL="false" color="automatic">
        <one:PageSize>
            <one:Automatic />
        </one:PageSize>
        <one:RuleLines visible="false" />
    </one:PageSettings>
    <one:Title lang="en-US">
        <one:OE author="Leonard Jonathan Oh" authorInitials="LJO" lastModifiedBy="Leonard Jonathan Oh" lastModifiedByInitials="LJO" creationTime="2021-08-06T15:36:49.000Z" lastModifiedTime="2021-08-06T15:36:57.000Z" objectID="{85905A1F-0185-439F-8BFF-44E9530204C5}{17}{B0}" alignment="left" quickStyleIndex="0">
            <one:T><![CDATA[p0.0 test]]></one:T>
        </one:OE>
    </one:Title>
    <one:Outline selected="all" author="Leonard Jonathan Oh" authorInitials="LJO" lastModifiedBy="Leonard Jonathan Oh" lastModifiedByInitials="LJO" lastModifiedTime="2021-08-06T16:08:25.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{10}{B0}">
        <one:Position x="36.0" y="86.4000015258789" z="0" />
        <one:Size width="250.4838562011719" height="344.3576049804687" />
        <one:OEChildren selected="partial">
            <one:OE creationTime="2021-08-06T15:47:58.000Z" lastModifiedTime="2021-08-06T15:47:58.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{76}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                <one:T selected="all"><![CDATA[My notebook:]]></one:T>
            </one:OE>
            <one:OE creationTime="2021-08-06T15:48:08.000Z" lastModifiedTime="2021-08-06T15:48:08.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{93}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                <one:T selected="all"><![CDATA[]]></one:T>
                <one:OEChildren selected="partial">
                    <one:OE creationTime="2021-08-06T15:47:01.000Z" lastModifiedTime="2021-08-06T15:47:53.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{11}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                        <one:List>
                            <one:Bullet bullet="2" fontSize="11.0" />
                        </one:List>
                        <one:T selected="all"><![CDATA[hello]]></one:T>
                    </one:OE>
                    <one:OE creationTime="2021-08-06T15:47:02.000Z" lastModifiedTime="2021-08-06T15:47:02.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{16}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                        <one:List>
                            <one:Bullet bullet="2" fontSize="11.0" />
                        </one:List>
                        <one:T selected="all"><![CDATA[world]]></one:T>
                    </one:OE>
                </one:OEChildren>
            </one:OE>
            <one:OE creationTime="2021-08-06T15:48:00.000Z" lastModifiedTime="2021-08-06T15:48:00.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{85}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                <one:T selected="all"><![CDATA[]]></one:T>
            </one:OE>
            <one:OE creationTime="2021-08-06T15:48:04.000Z" lastModifiedTime="2021-08-06T15:48:04.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{73}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                <one:T selected="all"><![CDATA[Thumbnails:]]></one:T>
            </one:OE>
            <one:OE creationTime="2021-08-06T15:48:07.000Z" lastModifiedTime="2021-08-06T15:48:07.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{90}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                <one:T selected="all"><![CDATA[]]></one:T>
            </one:OE>
            <one:OE creationTime="2021-08-06T15:47:14.000Z" lastModifiedTime="2021-08-06T15:47:14.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{51}{B0}" selected="all" alignment="left">
                <one:Image format="png" selected="all">
                    <one:Size width="19.5" height="25.875" isSetByUser="true" />
                    <one:Data>iVBORw0KGgoAAAANSUhEUgAAADQAAABFCAIAAACKZFLYAAAAAXNSR0IArs4c6QAAAARnQU1BAACx
jwv8YQUAAAAJcEhZcwAAHYcAAB2HAY/l8WUAAAOjSURBVGhD7ZpbSBRhFMfnorvufRW1rEwz1AdD
K7QUo7QMioqsHgr0KYgMoudeeu+xoKAQogeNKIKgKBO1e3SjEldyW+/obu7V3XHdy6y7nZyjaLq7
OW6zE8yPj9n/OfMN/Dmz3/lmmSXr6+sJsULhpyghj52o19apldvTaB2NOaEIeyO2Zif7M4TxEijt
fo2mViW8M4BSkVlnMjBYDkpVoUCZDMAfquWgYp9OLqJeEKIwt+HyGm5gPIdUOb5I5vgimeOLqM2R
5/vOoowC82La3cZgsEJ0BzSaGiUGURi7ODHf4UBzgiN+5Xg7A1ZzLSCK2woF4wbGc4jCnLR9JZr/
3By0A1QrZzXXAvH73L9mVX0uiUjm+CJqc6JYEKiWIIrKSdtXopHM8UUyxxfJHF9EbU7avuIRdfsy
v3Za2sL2DiX7sRAGCAgh6fcEcUryoLIV1VmZJXp9HkmlwAABISRnDIVjXZM4K0lEva1ymXKttsre
ofA5fJgSnDjfOb0+P2jInbYlxx+9t+H4wS07TlXUnqyoOVxaWZFfrFeqzW5HIMTC6ZJ1eZMBlhkJ
KAuivgRaJZ4Or7ZOPa85wUF3tj4oXpvL+H0G89CYy6ZTqLfmbq4pKnN6mfFJ+6VDjaNOK8OS5r5+
7aY0vCihxDBH+tlgy/uOzyM/MEEQULyGnfvkKam33radrj7AzoQu3L0e8E2lbhuQqWU4KXHE+t16
revhQmfAp2EjJEGARTim0ilwlCvU9q9Tv08LCNVvM6NcACQ/DPVB8TCehfTqUQkF2dRyBeUssAKa
9hxJoRa94TzXehWO3mmHbreVyySQFWxfveaRGy8fhcIzGC9AKdOhEgoyGGJvvnoMnjARnSDrV1YN
YZA4YlSOjEQisB7/xp9rcDTwZRADQaAnCtOfGj7ZGDcm5jhZXnO+9ihNUcaJMS5j6+mnvYI+DVAb
0rNQLmBjRvauwi2wSbz40Y0pgvCNJ341xGaZJlyeV9RYWQd9pPn1ky+jJi4ZCE6nlpr+URNGtQQS
bmimWmtxOwftFogLMnNydBn2Kc+dj53fLaPcJMB4r11HJd5ZbOhgWe5MJJyh0pSuL4C76fC63w70
3n737KfHhVMIwtrbL3cs2vWE4c8mvBSPxep6/k0uE7psQJznOWuPydn1NSnOgKjm/B6v8X572Die
JpdjSnBoTYDwM4x/yqfKSofY+n3QZRq2vOkOmsZVdHIKNo/0FzW+SOb4IpnjB0H8As43dCHTaPIF
AAAAAElFTkSuQmCC</one:Data>
                </one:Image>
                <one:OEChildren selected="partial">
                    <one:OE creationTime="2021-08-06T15:47:17.000Z" lastModifiedTime="2021-08-06T15:47:23.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{36}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                        <one:List>
                            <one:Bullet bullet="2" fontSize="11.0" />
                        </one:List>
                        <one:T selected="all"><![CDATA[this is a thumbail of my notebook]]></one:T>
                    </one:OE>
                </one:OEChildren>
            </one:OE>
            <one:OE creationTime="2021-08-06T15:47:34.000Z" lastModifiedTime="2021-08-06T15:47:34.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{58}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                <one:T selected="all"><![CDATA[]]></one:T>
            </one:OE>
            <one:OE creationTime="2021-08-06T15:47:35.000Z" lastModifiedTime="2021-08-06T15:47:35.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{66}{B0}" selected="all" alignment="left">
                <one:Image format="png" selected="all">
                    <one:Size width="24.0" height="27.75" isSetByUser="true" />
                    <one:Data>iVBORw0KGgoAAAANSUhEUgAAAEAAAABKCAIAAACEkMXvAAAAAXNSR0IArs4c6QAAAARnQU1BAACx jwv8YQUAAAAJcEhZcwAAHYcAAB2HAY/l8WUAAAGCSURBVGhD7ZYxS8NQEMdjaEGkICIizoUsUsTN WTqqCCIiqGDBRfwI4uDsqH4AdRIEwbWjg5uDCC4dHEWhdhA6+iCHQ+2lbUryf6f/HyG5CyTcL+/l vRuJoiiwTChXs1AADQXQUACNeQF1IwsLxcpGrby4NDY5Jbfyot1q1o8OPl8bkieijsDc5l5lfTf/ 6h2j4xPV41NJeqEKlKvLEiFwDhL1QhXo/xVYuAplydbtQ3xI3g2OABoKoKEAGvMCajOXvPo6nm8u Hi/PJRmQ+e392bUdSRSuVhd+anBxHPwm/Qikrt4xzLMdeD2F3IePD8m74bUAWwkL/GMBtxRKNDjD PNtB+n0gazLfBzyBAmjMC3j9E0uUiNcjEPcRyZ+S/wAaCqChABoKoDEvwFYiS9hKWIACaCiAhgJo zAuwlUDz16fQ+8vT/cnh18dbfNND1BFot5r+V+9QBRr1O3f2vHqHOoXCQrE0PbNydi157vS5CqkC VuAyioYCaCiAhgJojAsEwTdS82gV4OAExQAAAABJRU5ErkJggg==</one:Data>
                </one:Image>
                <one:OEChildren selected="partial">
                    <one:OE creationTime="2021-08-06T15:47:37.000Z" lastModifiedTime="2021-08-06T15:47:41.000Z" objectID="{2B41C8AC-8D4A-459D-B12C-0107571063C2}{67}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                        <one:List>
                            <one:Bullet bullet="2" fontSize="11.0" />
                        </one:List>
                        <one:T selected="all"><![CDATA[this is another thumbnail of my notebook]]></one:T>
                    </one:OE>
                </one:OEChildren>
            </one:OE>
            <one:OE creationTime="2021-08-06T16:07:37.000Z" lastModifiedTime="2021-08-06T16:07:37.000Z" objectID="{902AD630-91B7-4D8B-96A7-107572724072}{28}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                <one:T selected="all"><![CDATA[]]></one:T>
            </one:OE>
            <one:OE creationTime="2021-08-06T16:07:42.000Z" lastModifiedTime="2021-08-06T16:07:42.000Z" objectID="{902AD630-91B7-4D8B-96A7-107572724072}{10}{B0}" selected="all" alignment="left" quickStyleIndex="1">
                <one:T selected="all"><![CDATA[Attachments:]]></one:T>
            </one:OE>
            <one:OE creationTime="2021-08-06T16:08:25.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" objectID="{902AD630-91B7-4D8B-96A7-107572724072}{49}{B0}" selected="all" alignment="left">
                <one:InsertedFile selected="all" pathCache="C:\Users\LeonardJonathan\AppData\Local\Microsoft\OneNote\16.0\cache\00001BNS.bin" pathSource="C:\Users\LeonardJonathan\Desktop\attachment1(something in brackets).txt" preferredName="attachment1(something in brackets).txt" />
            </one:OE>
            <one:OE creationTime="2021-08-06T16:08:10.000Z" lastModifiedTime="2021-08-06T16:08:10.000Z" objectID="{902AD630-91B7-4D8B-96A7-107572724072}{41}{B0}" selected="all" alignment="left">
                <one:InsertedFile selected="all" pathCache="C:\Users\LeonardJonathan\AppData\Local\Microsoft\OneNote\16.0\cache\00001BNS.bin" pathSource="C:\Users\LeonardJonathan\Desktop\attachment2[something in brackets].txt" preferredName="attachment2(something in brackets).txt" />
            </one:OE>
        </one:OEChildren>
    </one:Outline>
</one:Page>
'@ -as [xml]
    $page
}

Describe 'New-SectionGroupConversionConfig' -Tag 'Unit' {

    Context 'Behavior' {

        BeforeEach {
            function Get-OneNotePageContent {}
            Mock Get-OneNotePageContent {
                Get-FakeOneNotePageContent
            }
            $fakeHierarchy = Get-FakeOneNoteHierarchy
            $params = @{
                OneNoteConnection = 'some connection'
                NotesDestination = 'c:\temp\notes'
                Config = Get-DefaultConfiguration
                SectionGroups = $fakeHierarchy.Notebooks.Notebook
                LevelsFromRoot = 0
                AsArray = $false
            }
        }

        It "Should ignore empty Section Groups and Section(s)" {
            $fakeHierarchy = Get-FakeOneNoteHierarchyWithEmptySectionGroupsAndSectionsAndPages
            $params['SectionGroups'] = $fakeHierarchy.Notebooks.Notebook
            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            $result.Count | Should -Be 0
        }

        It "Should construct individual conversion configuration(s) for pages, based on a given Section Group XML object. Ignores pages in recycle bin." {
            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48
        }

        It "Should get its pagePrefix from its parent page (if any)" {
            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            $pagePrefixSeparatorChar = [io.path]::DirectorySeparatorChar
            for ($i = 0; $i -lt $result.Count; $i = $i + 8) { # Test in eights
                $pageCfg1 = $result[$i] # First level page
                $pageCfg2 = $result[$i + 1] # Second level page preceded by a first level page
                $pageCfg3 = $result[$i + 2] # Third level page preceded by a second level page
                $pageCfg4 = $result[$i + 3] # Second level page preceded by a third level page
                $pageCfg5 = $result[$i + 4] # First level page preceded by a second level page
                $pageCfg6 = $result[$i + 5] # Third level page preceded by a first level page
                $pageCfg7 = $result[$i + 6] # Second level page preceded by a third level page
                $pageCfg8 = $result[$i + 7] # First level page preceded by a second level page

                $pageCfg1['pagePrefix'] | Should -Be ''
                $pageCfg2['pagePrefix'] | Should -Be "$( $pageCfg1['filePathRel'] )$pagePrefixSeparatorChar"
                $pageCfg3['pagePrefix'] | Should -Be "$( $pageCfg2['filePathRel'] )$pagePrefixSeparatorChar"
                $pageCfg4['pagePrefix'] | Should -Be "$( $pageCfg1['filePathRel'] )$pagePrefixSeparatorChar"

                $pageCfg5['pagePrefix'] | Should -Be '' # First level page
                $pageCfg6['pagePrefix'] | Should -Be "$( $pageCfg5['filePathRel'] )$pagePrefixSeparatorChar" # Third level page preceded by a first level page
                $pageCfg7['pagePrefix'] | Should -Be "$( $pageCfg5['filePathRel'] )$pagePrefixSeparatorChar" # Second level page preceded by a third level page
                $pageCfg8['pagePrefix'] | Should -Be '' # First level page
            }
        }

        It "Should determine file and folder paths correctly" {
            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                $pageCfg['fileName'] | Should -Match "$( [regex]::Escape($pageCfg['fileExtension']) )$"
                $pageCfg['filePathNormal'] | Should -Match  "$( [regex]::Escape($pageCfg['fileName']) )$"
                $pageCfg['filePathLong'] | Should -Match "$( [regex]::Escape($pageCfg['fileName']) )$"
                $pageCfg['filePath'] | Should -Match "$( [regex]::Escape($pageCfg['fileName']) )$"
            }
        }

        It "Should determine attachment references correctly" {
            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                $pageCfg['insertedAttachments'].Count | Should -Be 2

                $pageCfg['insertedAttachments'][0]['markdownFileName'] | Should -Be 'attachment1\(something-in-brackets\).txt'
                $pageCfg['insertedAttachments'][1]['markdownFileName'] | Should -Be 'attachment2\(something-in-brackets\).txt'
            }
        }

        It "Should generate fully unique docx file for each page, even for identically-named pages in a section" {
            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            for ($i = 0; $i -lt $result.Count; $i++) {
                $count = 0
                for ($j = 0; $j -lt $result.Count; $j++) {
                    if ($result[$i] -eq $result[$j]) {
                        $count++
                    }
                }
                $count | Should -Be 1
            }

            $fakeHierarchy = Get-FakeOneNoteHierarchyWithDuplicatePageNames  # Duplicate page names within sections
            $params['SectionGroups'] = $fakeHierarchy.Notebooks.Notebook
            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 18

            for ($i = 0; $i -lt $result.Count; $i++) {
                $count = 0
                for ($j = 0; $j -lt $result.Count; $j++) {
                    if ($result[$i] -eq $result[$j]) {
                        $count++
                    }
                }
                $count | Should -Be 1
            }
        }

        It "Should append a unique postfix to identically named pages of a Section" {
            $fakeHierarchy = Get-FakeOneNoteHierarchyWithDuplicatePageNames  # Duplicate page names within sections
            $params['SectionGroups'] = $fakeHierarchy.Notebooks.Notebook
            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 18

            for ($i = 0; $i -lt $result.Count; $i = $i + 3) { # Test in threes
                $pageCfg1 = $result[$i]
                $pageCfg2 = $result[$i + 1]
                $pageCfg3 = $result[$i + 2]

                $pageCfg2['filePathRel'] | Should -Be "$( $pageCfg1['filePathRel'] )-1"
                $pageCfg3['filePathRel'] | Should -Be "$( $pageCfg1['filePathRel'] )-2"
            }
        }

        It "Should honor config notesdestpath" {
            if ($env:OS -match 'windows') {
                $params['Config']['notesdestpath']['value'] = 'D:\foo\bar'
                $params['NotesDestination'] = 'D:\foo\bar'
            }else {
                $params['Config']['notesdestpath']['value'] = '/foo/bar'
                $params['NotesDestination'] = '/foo/bar'
            }

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            # Test the first object
            $pageCfg = $result[0]
            $regex = "^$( [regex]::Escape($params['Config']['notesdestpath']['value']) )"
            $regexTmp = "^$( [regex]::Escape($pageCfg['tmpPath']) )"
            $pageCfg['fileDirectory'] | Should -Match $regex
            $pageCfg['filePathNormal'] | Should -Match $regex
            $pageCfg['mediaParentPath'] | Should -Match $regex
            $pageCfg['mediaPath'] | Should -Match $regex
            $pageCfg['mediaParentPathPandoc'] | Should -Be $pageCfg['tmpPath'].Replace([io.path]::DirectorySeparatorChar, '/')
            $pageCfg['mediaPathPandoc'] | Should -Be $( [io.path]::combine($pageCfg['tmpPath'], 'media').Replace([io.path]::DirectorySeparatorChar, '/') )
            $pageCfg['docxExportFilePath'] | Should -Match $regex
            $pageCfg['insertedAttachments'] | ForEach-Object {
                $_['destination'] | Should -Match $regex
            }
            $pageCfg['directoriesToCreate'] | ForEach-Object {
                ($_ -match $regex) -or ($_ -match $regexTmp) | Should -Be $true
            }
            $pageCfg['directoriesToDelete'] | ForEach-Object {
                $_ | Should -Match $regexTmp
            }
        }

        It "Should honor config prefixFolders" {
            $params['Config']['prefixFolders']['value'] = 1

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            for ($i = 0; $i -lt $result.Count; $i = $i + 8) { # Test in eights
                $pageCfg1 = $result[$i] # First level page
                $pageCfg2 = $result[$i + 1] # Second level page preceded by a first level page
                $pageCfg3 = $result[$i + 2] # Third level page preceded by a second level page
                $pageCfg4 = $result[$i + 3] # Second level page preceded by a third level page
                $pageCfg5 = $result[$i + 4] # First level page preceded by a second level page
                $pageCfg6 = $result[$i + 5] # Third level page preceded by a first level page
                $pageCfg7 = $result[$i + 6] # Second level page preceded by a third level page
                $pageCfg8 = $result[$i + 7] # First level page preceded by a second level page

                Split-Path $pageCfg1['filePathNormal'] -Leaf | Should -Be  $pageCfg1['fileName']
                Split-Path $pageCfg1['filePathNormal'] -Parent | Should -Be $pageCfg1['fileDirectory']
                $pageCfg1['levelsPrefix']| Should -Be "$( '../' * ($pageCfg1['levelsFromRoot'] + $pageCfg1['pageLevel'] - 1) )"

                Split-Path $pageCfg2['filePathNormal'] -Leaf | Should -Be  $pageCfg2['fileName']
                Split-Path $pageCfg2['filePathNormal'] -Parent | Should -Be $pageCfg2['fileDirectory']
                $pageCfg2['levelsPrefix']| Should -Be "$( '../' * ($pageCfg2['levelsFromRoot'] + $pageCfg2['pageLevel'] - 1) )"

                Split-Path $pageCfg3['filePathNormal'] -Leaf | Should -Be  $pageCfg3['fileName']
                Split-Path $pageCfg3['filePathNormal'] -Parent | Should -Be $pageCfg3['fileDirectory']
                $pageCfg3['levelsPrefix']| Should -Be "$( '../' * ($pageCfg3['levelsFromRoot'] + $pageCfg3['pageLevel'] - 1) )"

                Split-Path $pageCfg4['filePathNormal'] -Leaf | Should -Be $pageCfg4['fileName']
                Split-Path $pageCfg4['filePathNormal'] -Parent | Should -Be $pageCfg4['fileDirectory']
                $pageCfg4['levelsPrefix']| Should -Be "$( '../' * ($pageCfg4['levelsFromRoot'] + $pageCfg4['pageLevel'] - 1) )"

                Split-Path $pageCfg5['filePathNormal'] -Leaf | Should -Be $pageCfg5['fileName']
                Split-Path $pageCfg5['filePathNormal'] -Parent | Should -Be $pageCfg5['fileDirectory']
                $pageCfg5['levelsPrefix']| Should -Be "$( '../' * ($pageCfg5['levelsFromRoot'] + $pageCfg5['pageLevel'] - 1) )"

                Split-Path $pageCfg6['filePathNormal'] -Leaf | Should -Be $pageCfg6['fileName']
                Split-Path $pageCfg6['filePathNormal'] -Parent | Should -Be $pageCfg6['fileDirectory']
                $pageCfg6['levelsPrefix']| Should -Be "$( '../' * ($pageCfg6['levelsFromRoot'] + $pageCfg6['pageLevel'] - 1) )"

                Split-Path $pageCfg7['filePathNormal'] -Leaf | Should -Be $pageCfg7['fileName']
                Split-Path $pageCfg7['filePathNormal'] -Parent | Should -Be $pageCfg7['fileDirectory']
                $pageCfg7['levelsPrefix']| Should -Be "$( '../' * ($pageCfg7['levelsFromRoot'] + $pageCfg7['pageLevel'] - 1) )"

                # Test the first level page preceded by a second level page
                Split-Path $pageCfg8['filePathNormal'] -Leaf | Should -Be $pageCfg8['fileName']
                Split-Path $pageCfg8['filePathNormal'] -Parent | Should -Be $pageCfg8['fileDirectory']
                $pageCfg8['levelsPrefix']| Should -Be "$( '../' * ($pageCfg8['levelsFromRoot'] + $pageCfg8['pageLevel'] - 1) )"
            }

            $params['Config']['prefixFolders']['value'] = 2

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            for ($i = 0; $i -lt $result.Count; $i = $i + 8) { # Test in eights
                $pageCfg1 = $result[$i] # First level page
                $pageCfg2 = $result[$i + 1] # Second level page preceded by a first level page
                $pageCfg3 = $result[$i + 2] # Third level page preceded by a second level page
                $pageCfg4 = $result[$i + 3] # Second level page preceded by a third level page
                $pageCfg5 = $result[$i + 4] # First level page preceded by a second level page
                $pageCfg6 = $result[$i + 5] # Third level page preceded by a first level page
                $pageCfg7 = $result[$i + 6] # Second level page preceded by a third level page
                $pageCfg8 = $result[$i + 7] # First level page preceded by a second level page

                Split-Path $pageCfg1['filePathNormal'] -Parent | Should -Be $pageCfg1['fileDirectory']
                $pageCfg1['filePathRelUnderscore'] | Should -Be "$( $pageCfg1['nameCompat'] )"
                Split-Path $pageCfg1['filePathNormal'] -Leaf | Should -Be "$( $pageCfg1['filePathRelUnderscore'] ).md"
                $pageCfg1['levelsPrefix']| Should -Be "$( '../' * ($pageCfg1['levelsFromRoot'] + 1 - 1) )"

                Split-Path $pageCfg2['filePathNormal'] -Parent | Should -Be $pageCfg2['fileDirectory']
                $pageCfg2['filePathRelUnderscore'] | Should -Be "$( $pageCfg1['nameCompat'] )_$( $pageCfg2['nameCompat'] )"
                Split-Path $pageCfg2['filePathNormal'] -Leaf | Should -Be "$( $pageCfg2['filePathRelUnderscore'] ).md"
                $pageCfg2['levelsPrefix']| Should -Be "$( '../' * ($pageCfg2['levelsFromRoot'] + 1 - 1) )"

                Split-Path $pageCfg3['filePathNormal'] -Parent | Should -Be $pageCfg3['fileDirectory']
                $pageCfg3['filePathRelUnderscore'] | Should -Be "$( $pageCfg1['nameCompat'] )_$( $pageCfg2['nameCompat'] )_$( $pageCfg3['nameCompat'] )"
                Split-Path $pageCfg3['filePathNormal'] -Leaf | Should -Be "$( $pageCfg3['filePathRelUnderscore'] ).md"
                $pageCfg3['levelsPrefix']| Should -Be "$( '../' * ($pageCfg3['levelsFromRoot'] + 1 - 1) )"

                Split-Path $pageCfg4['filePathNormal'] -Parent | Should -Be $pageCfg4['fileDirectory']
                $pageCfg4['filePathRelUnderscore'] | Should -Be "$( $pageCfg1['nameCompat'] )_$( $pageCfg4['nameCompat'] )"
                Split-Path $pageCfg4['filePathNormal'] -Leaf | Should -Be "$( $pageCfg4['filePathRelUnderscore'] ).md"
                $pageCfg4['levelsPrefix']| Should -Be "$( '../' * ($pageCfg4['levelsFromRoot'] + 1 - 1) )"

                Split-Path $pageCfg5['filePathNormal'] -Parent | Should -Be $pageCfg5['fileDirectory']
                $pageCfg5['filePathRelUnderscore'] | Should -Be "$( $pageCfg5['nameCompat'] )"
                Split-Path $pageCfg5['filePathNormal'] -Leaf | Should -Be "$( $pageCfg5['filePathRelUnderscore'] ).md"
                $pageCfg5['levelsPrefix']| Should -Be "$( '../' * ($pageCfg5['levelsFromRoot'] + 1 - 1) )"

                Split-Path $pageCfg6['filePathNormal'] -Parent | Should -Be $pageCfg6['fileDirectory']
                $pageCfg6['filePathRelUnderscore'] | Should -Be "$( $pageCfg5['nameCompat'] )_$( $pageCfg6['nameCompat'] )"
                Split-Path $pageCfg6['filePathNormal'] -Leaf | Should -Be "$( $pageCfg6['filePathRelUnderscore'] ).md"
                $pageCfg6['levelsPrefix']| Should -Be "$( '../' * ($pageCfg6['levelsFromRoot'] + 1 - 1) )"

                Split-Path $pageCfg7['filePathNormal'] -Parent | Should -Be $pageCfg7['fileDirectory']
                $pageCfg7['filePathRelUnderscore'] | Should -Be "$( $pageCfg5['nameCompat'] )_$( $pageCfg7['nameCompat'] )"
                Split-Path $pageCfg7['filePathNormal'] -Leaf | Should -Be "$( $pageCfg7['filePathRelUnderscore'] ).md"
                $pageCfg7['levelsPrefix']| Should -Be "$( '../' * ($pageCfg7['levelsFromRoot'] + 1 - 1) )"

                Split-Path $pageCfg8['filePathNormal'] -Parent | Should -Be $pageCfg8['fileDirectory']
                $pageCfg8['filePathRelUnderscore'] | Should -Be "$( $pageCfg8['nameCompat'] )"
                Split-Path $pageCfg8['filePathNormal'] -Leaf | Should -Be "$( $pageCfg8['filePathRelUnderscore'] ).md"
                $pageCfg8['levelsPrefix']| Should -Be "$( '../' * ($pageCfg8['levelsFromRoot'] + 1 - 1) )"
            }
        }

        It "Should honor config docxNamingConvention" {
            $params['Config']['docxNamingConvention']['value'] = 1

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                Split-Path $pageCfg['docxExportFilePath'] -Leaf | Should -Be "$( $pageCfg['id'] )-$( $pageCfg['lastModifiedTimeEpoch'] ).docx"
            }

            $params['Config']['docxNamingConvention']['value'] = 2

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                Split-Path $pageCfg['docxExportFilePath'] -Leaf | Should -Be "$( $pageCfg['pathFromRootCompat'] ).docx"
            }
        }

        It "Should honor config mdFileNameAndFolderNameMaxLength, and normalize the final .md file and its parent folder names to a maximum of x characters" {
            # 5 notes, each with very long names
            $fakeHierarchy = @'
<?xml version="1.0"?>
<one:Notebooks xmlns:one="http://schemas.microsoft.co1m/office/onenote/2013/onenote">
    <one:Notebook name="test" nickname="test" ID="{38E47DAB-211E-4EC1-85F1-129656A9D2CE}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/" lastModifiedTime="2021-08-06T16:27:58.000Z" color="#ADE792">
        <one:Section name="s0" ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/s0.one" lastModifiedTime="2021-08-06T16:08:25.000Z" color="#8AA8E4">
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096991}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19542261697052950701320178013485171541838441}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19535140647270019211520151454305551340000401}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:46:36.000Z" pageLevel="3" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19542261697052950701320178013485171541838442}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096992}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
        </one:Section>
    </one:Notebook>
</one:Notebooks>
'@ -as [xml]
            $params['SectionGroups'] = $fakeHierarchy.Notebooks.Notebook
            $params['Config']['mdFileNameAndFolderNameMaxLength']['value'] = 100

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 5 pages from 'test' notebook
            $result.Count | Should -Be 5

            foreach ($pageCfg in $result) {
                # The names of this `.md` file and all of its parent folders should be no longer than x characters
                $pageCfg['fileName'].Length | Should -Be $params['Config']['mdFileNameAndFolderNameMaxLength']['value']

                $split = $pageCfg['filePath'].Split([io.path]::DirectorySeparatorChar)
                foreach ($s in $split) {
                    $s.Length | Should -BeLessOrEqual $params['Config']['mdFileNameAndFolderNameMaxLength']['value']
                }
            }

        }

        It "Should honor config medialocation" {
            $params['Config']['medialocation']['value'] = 1

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                $pageCfg['mediaParentPath'] | Should -Be $pageCfg['notesBaseDirectory']
            }

            $params['Config']['medialocation']['value'] = 2

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                $pageCfg['mediaParentPath'] | Should -Be $pageCfg['fileDirectory']
            }
        }

        It "Should honor config conversion" {
            $params['Config']['conversion']['value'] = 'gfm+pipe_tables'

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                $pageCfg['conversion'] | Should -Be $params['Config']['conversion']['value']
            }
        }

        It "Should honor config headerTimestampEnabled" {
            $params['Config']['headerTimestampEnabled']['value'] = 1

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                # The first line will be replaced by page header (8 lines)
                $fakeMarkdownContent = ''

                # Mutate
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        $fakeMarkdownContent = $fakeMarkdownContent -replace $r['searchRegex'], $r['replacement']
                    }
                }

                # Should add only a title, newline, date, newline, separator
                $fakeMarkdownContent = $fakeMarkdownContent -split "`n"
                $fakeMarkdownContent.Count | Should -Be 8
                $fakeMarkdownContent[0] | Should -Match '^# .*$'
                $fakeMarkdownContent[1] | Should -Match '^\s*$'
                $fakeMarkdownContent[2] -match '^Created: (.+)$' | Should -Be $true
                if ($fakeMarkdownContent[2] -match '^Created: (.+)$') {
                    Get-Date $matches[1] | Should -BeOfType [Datetime]
                }
                $fakeMarkdownContent[3] | Should -Match '^\s*$'
                $fakeMarkdownContent[4] -match '^Modified: (.+)$' | Should -Be $true
                if ($fakeMarkdownContent[4] -match '^Modified: (.+)$') {
                    Get-Date $matches[1] | Should -BeOfType [Datetime]
                }
                $fakeMarkdownContent[5] | Should -Match '^\s*$'
                $fakeMarkdownContent[6] | Should -Be '---'
                $fakeMarkdownContent[7] | Should -Match '^\s*$'
            }

            $params['Config']['headerTimestampEnabled']['value'] = 2

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                $fakeMarkdownContent = ''

                # Mutate
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        $fakeMarkdownContent = $fakeMarkdownContent -replace $r['searchRegex'], $r['replacement']
                    }
                }

                # Should add only a title
                $fakeMarkdownContent = $fakeMarkdownContent -split "`n"
                $fakeMarkdownContent.Count | Should -Be 1
                $fakeMarkdownContent[0] | Should -Match '^# .*'
            }
        }

        It "Should honor config keepspaces" {
            $params['Config']['keepspaces']['value'] = 1

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            # The first line will be replaced by page header (8 lines)
            $fakeMarkdownContent = @"

hello world$( [char]0x00A0 )
- foo

  - foo1

- bar

  - bar1

>
>
> some other text
"@ -replace "`r", '' # On some Windows Powershell 5 versions, a here-string will contain `\r`, so let's ensure that doesn't happen.

            foreach ($pageCfg in $result) {
                # Mutate
                $mutated = $fakeMarkdownContent
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        $mutated = $mutated -replace $r['searchRegex'], $r['replacement']
                    }
                }

                # Should remove newlines between bullets, and remove non-breaking spaces. Ignore first 8 lines for page header
                $split = $mutated -split "`n"
                $expectedBody = $split[8..($split.Count - 1)] -join "`n"
                $expectedBody | Should -Be $( @"
hello world
- foo
  - foo1
- bar
  - bar1



some other text
"@ -replace "`r", '') # On some Windows Powershell 5 versions, a here-string will contain `\r`, so let's ensure that doesn't happen.
            }

            $params['Config']['keepspaces']['value'] = 2

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                # Mutate
                $mutated = $fakeMarkdownContent
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        $mutated = $mutated -replace $r['searchRegex'], $r['replacement']
                    }
                }

                # Should keep newlines between bullets, and keep non-breaking spaces. Ignore first 8 lines for page header
                $split = $mutated -split "`n"
                $expectedBody = $split[8..($split.Count - 1)] -join "`n"
                $expectedBody | Should -Be $( @"
hello world$( [char]0x00A0 )
- foo

  - foo1

- bar

  - bar1

>
>
> some other text
"@ -replace "`r", '') # On some Windows Powershell 5 versions, a here-string will contain `\r`, so let's ensure that doesn't happen.
            }

        }

        It "Should honor config keepescape" {
            $params['Config']['keepescape']['value'] = 1

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                # The first line will be replaced by page header (8 lines)
                $fakeMarkdownContent = @"

hello\$\^\\\*\_\[\]world
foo\bar\0
"@

                # Mutate
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        $fakeMarkdownContent = $fakeMarkdownContent -replace $r['searchRegex'], $r['replacement']
                    }
                }

                # Should remove all backslashes. Ignore first 8 lines for page header
                $fakeMarkdownContent = $fakeMarkdownContent -split "`n"
                $fakeMarkdownContent.Count | Should -Be 10
                $fakeMarkdownContent[8] | Should -Be 'hello$^*_[]world'
                $fakeMarkdownContent[9] | Should -Be 'foobar0'
            }
            $params['Config']['keepescape']['value'] = 2

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                # The first line will be replaced by page header (8 lines)
                $fakeMarkdownContent = @'

hello\$\^\\\*\_\[\]world
foo\bar\0
'@

                # Mutate
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        $fakeMarkdownContent = $fakeMarkdownContent -replace $r['searchRegex'], $r['replacement']
                    }
                }

                # Should remove backslashes that precede non-alphanumeric characters. Ignore first 8 lines for page header
                $fakeMarkdownContent = $fakeMarkdownContent -split "`n"
                $fakeMarkdownContent.Count | Should -Be 10
                $fakeMarkdownContent[8] | Should -Be 'hello$^\*_[]world'
                $fakeMarkdownContent[9] | Should -Be 'foo\bar\0'
            }

            $params['Config']['keepescape']['value'] = 3

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                # The first line will be replaced by page header (8 lines)
                $fakeMarkdownContent = @'

hello\$\^\\\*\_\[\]world
foo\bar\0
'@

                # Mutate
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        $fakeMarkdownContent = $fakeMarkdownContent -replace $r['searchRegex'], $r['replacement']
                    }
                }

                # Should keep backslashes. Ignore first 8 lines for page header
                $fakeMarkdownContent = $fakeMarkdownContent -split "`n"
                $fakeMarkdownContent.Count | Should -Be 10
                $fakeMarkdownContent[8] | Should -Be 'hello\$\^\\\*\_\[\]world'
                $fakeMarkdownContent[9] | Should -Be 'foo\bar\0'
            }
        }

        It "Should honor config newlineCharacter" {
            $params['Config']['newlineCharacter']['value'] = 1

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                # The first line will be replaced by page header (8 lines)
                $fakeMarkdownContent = @"

foo`r`nbar`r`nbaz
"@

                # Mutate
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        $fakeMarkdownContent = $fakeMarkdownContent -replace $r['searchRegex'], $r['replacement']
                    }
                }

                # Should remove CRs. Ignore first 8 lines for page header
                $fakeMarkdownContent = $fakeMarkdownContent -split "`n"
                $fakeMarkdownContent.Count | Should -Be 11
                $fakeMarkdownContent[8] | Should -Match '^foo$'
                $fakeMarkdownContent[9] | Should -Match '^bar$'
                $fakeMarkdownContent[10] | Should -Match '^baz$'`
            }

            $params['Config']['newlineCharacter']['value'] = 2

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                # The first line will be replaced by page header (8 lines)
               $fakeMarkdownContent = @"

foo`r`nbar`r`nbaz
"@
                # Mutate
                foreach ($m in $pageCfg['mutations']) {
                    foreach ($r in $m['replacements']) {
                        $fakeMarkdownContent = $fakeMarkdownContent -replace $r['searchRegex'], $r['replacement']
                    }
                }

                # Should retain CRs. Ignore first 8 lines for page header
                $fakeMarkdownContent = $fakeMarkdownContent -split "`n"
                $fakeMarkdownContent.Count | Should -Be 11
                $fakeMarkdownContent[8] | Should -Match "^foo`r$"
                $fakeMarkdownContent[9] | Should -Match "^bar`r$"
                $fakeMarkdownContent[10] | Should -Match "^baz$"
            }
        }

        It "Should honor config exportPdf" {
            $params['Config']['exportPdf']['value'] = 2

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 15 pages from 'test' notebook, 15 pages from 'test2' notebook
            $result.Count | Should -Be 48

            foreach ($pageCfg in $result) {
                $pageCfg['pdfExportFilePath'] | Should -Be ($pageCfg['FilePath'] -replace '\.md$', '.pdf')
            }

            # 5 notes, each with very long names
            $fakeHierarchy = @'
<?xml version="1.0"?>
<one:Notebooks xmlns:one="http://schemas.microsoft.co1m/office/onenote/2013/onenote">
    <one:Notebook name="test" nickname="test" ID="{38E47DAB-211E-4EC1-85F1-129656A9D2CE}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/" lastModifiedTime="2021-08-06T16:27:58.000Z" color="#ADE792">
        <one:Section name="s0" ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{B0}" path="https://d.docs.live.net/741e69cc14cf9571/Skydrive Notebooks/test/s0.one" lastModifiedTime="2021-08-06T16:08:25.000Z" color="#8AA8E4">
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096991}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19542261697052950701320178013485171541838441}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19535140647270019211520151454305551340000401}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:38:03.000Z" lastModifiedTime="2021-08-06T15:46:36.000Z" pageLevel="3" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19542261697052950701320178013485171541838442}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:36:14.000Z" lastModifiedTime="2021-08-06T15:38:01.000Z" pageLevel="2" />
            <one:Page ID="{3D017C7D-F890-4AC8-A094-DEC1163E7B85}{1}{E19461971475288592555920101886406896686096992}" name="Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name Some long name" dateTime="2021-08-06T15:36:33.000Z" lastModifiedTime="2021-08-06T16:08:25.000Z" pageLevel="1" />
        </one:Section>
    </one:Notebook>
</one:Notebooks>
'@ -as [xml]
            $params['SectionGroups'] = $fakeHierarchy.Notebooks.Notebook
            $params['Config']['mdFileNameAndFolderNameMaxLength']['value'] = 100
            $params['Config']['exportPdf']['value'] = 2

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # 5 pages from 'test' notebook
            $result.Count | Should -Be 5

            foreach ($pageCfg in $result) {
                $pageCfg['pdfExportFilePath'] | Should -Be ($pageCfg['FilePath'] -replace '.\.md$', '.pdf') # 1 character from the basename should have been trimmed when replacing the extension with '.pdf'
            }
        }

        It "-AsArray should construct a full Section Group conversion configuration object, based on a given Section Group XML object. Ignores pages in recycle bin." {
            $params['AsArray'] = $true

            $result = @( New-SectionGroupConversionConfig @params 6>$null )

            # Validate a Section Group conversion configuration recursively against its corresponding section group XML object
            function Assert-SectionGroupConversionConfigReflectsSectionGroupXmlObject ([array]$conversionCfg, [array]$sectionGroups) {
                # Treat a notebook as a Section Group. It is no different.

                # Skip recycle bin Section Groups
                $sectionGroups = @(
                    $sectionGroups | Where-Object { ! (Get-Member -InputObject $_ -Name 'isRecycleBin') }
                )

                # Should have equal number of Section Grouops
                $conversionCfg.Count | Should -Be $sectionGroups.Count

                for ($i = 0; $i -lt ($sectionGroups).Count; $i++) {
                    if (Get-Member -InputObject $sectionGroups[$i] -Name 'Section' -Membertype Properties) {
                        $sections = @(
                            $sectionGroups[$i].Section
                        )

                        # Should have equal number of sections
                        $conversionCfg[$i]['sections'].Count | Should -Be $sections.Count

                        for ($s = 0; $s -lt $sections.Count; $s++) {
                            if (Get-Member -InputObject $sections[$s] -Name 'Page' -Membertype Properties) {
                                $pages =@(
                                     $sections[$s].Page
                                )

                                # Should have equal number of pages
                                $conversionCfg[$i]['sections'][$s]['pages'].Count | Should -Be $pages.Count
                            }
                        }
                    }
                    if (Get-Member -InputObject $sectionGroups[$i] -Name 'SectionGroup' -Membertype Properties) {
                        $sectionGroups = @(
                            $sectionGroups[$i].SectionGroup
                        )
                        Assert-SectionGroupConversionConfigReflectsSectionGroupXmlObject $conversionCfg[$i]['sectiongroups'] $sectionGroups
                    }
                }
            }
            Assert-SectionGroupConversionConfigReflectsSectionGroupXmlObject $result[0] @( $fakeHierarchy.Notebooks.Notebook )
        }

    }

}

Describe 'Convert-OneNotePage' -Tag 'Unit' {
    function New-FakeSectionGroupConversionConfig {
        function Get-OneNotePageContent {}
        Mock Get-OneNotePageContent {
            Get-FakeOneNotePageContent
        }
        $params = @{
            OneNoteConnection = 'some connection'
            NotesDestination = 'c:\temp\notes'
            Config = Get-DefaultConfiguration
            SectionGroups = (Get-FakeOneNoteHierarchy).Notebooks.Notebook
            LevelsFromRoot = 0
        }
        $result = @( New-SectionGroupConversionConfig @params 6>$null )

        $result
    }

    Context 'Parameters' {
    }

    Context 'Behavior' {

        BeforeEach {
            Mock New-Item -ParameterFilter { $Path -and $ItemType -eq 'Directory' -and $Force} {
                [pscustomobject]@{
                    FullName = 'foo'
                }
            }
            Mock Remove-Item {}
            Mock Remove-Item -ParameterFilter { $LiteralPath -and $Force } {}
            function Publish-OneNotePage {
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
            }
            Mock Publish-OneNotePage {}
            Mock Publish-OneNotePage -ParameterFilter { $PublishFormat -eq 'pfWord' } {}
            Mock Publish-OneNotePage -ParameterFilter { $PublishFormat -eq 'pfPdf' } {}

            Mock Start-Process {
                [pscustomobject]@{
                    ExitCode = 0
                }
            }
            Mock Test-Path { $false }
            Mock Test-Path -ParameterFilter { $LiteralPath } { $false }
            Mock Copy-Item {}
            Mock Get-ChildItem {
                [pscustomobject]@{
                    BaseName = 'image'
                    Name = 'image1.jpg'
                    FullName = 'c:\temp\notes\mynotebook\media\image1.jpg'
                    Extension = '.jpg'
                }
            }
            Mock Move-Item {
                [pscustomobject]@{
                    BaseName = 'somenewname'
                    Name = 'somenewname.jpg'
                    FullName = 'c:\temp\notes\mynotebook\media\somenewname.jpg'
                    Extension = '.jpg'
                }
            }
            function Get-Content { # Create this function for the sake of a "bug" in pester where calls to the Get-Content mock spits out a non-terminating pester error: A parameter cannot be found that matches parameter name 'Raw'
                param (
                    [Parameter()]
                    [string]
                    $Path
                ,
                    [Parameter()]
                    [string]
                    $LiteralPath
                ,
                    [Parameter()]
                    [switch]
                    $Raw
                )
                ''
            }
            Mock Get-Content {
                ''
            }
            Mock Get-Content -ParameterFilter { $LiteralPath -and $Raw } {
                ''
            }
            Mock Get-Content -ParameterFilter { $LiteralPath -and ! $Raw } {
                ''
            }
            function Set-ContentNoBom {
                param (
                    [Parameter()]
                    [string]
                    $LiteralPath
                ,
                    [Parameter()]
                    [string]
                    $Value
                )
            }
            Mock Set-ContentNoBom {}
            $params = @{
                OneNoteConnection = 'some connection'
                Config = Get-DefaultConfiguration
                ConversionConfig = New-FakeSectionGroupConversionConfig | Select-Object -First 1
            }
        }

        It "Creates directories" {
            Convert-OneNotePage @params 6>$null

            Assert-MockCalled -CommandName New-Item -ParameterFilter { $Path -and $Force } -Times 5 -Scope It
        }

        It "Halts converting if creation of any directory fails" {
            Mock New-Item -ParameterFilter { $ItemType -eq 'Directory' -and $Force } { throw }

            $err = Convert-OneNotePage @params 6>$null 2>&1

            $err.Exception.Message | Select-Object -First 1 | Should -match 'Failed to convert page'
        }

        It "Removes existing docx by default" {
            Mock Test-Path -ParameterFilter { $LiteralPath } { $true }

            Convert-OneNotePage @params 6>$null

            Assert-MockCalled -CommandName Remove-Item -ParameterFilter { $LiteralPath -and $Force } -Times 2 -Scope It
        }

        It "Halts converting if removal of existing docx fails" {
            Mock Test-Path -ParameterFilter { $LiteralPath } { $true }
            Mock Remove-Item -ParameterFilter { $LiteralPath -and $Force } { throw }

            $err = Convert-OneNotePage @params 6>$null 2>&1

            Assert-MockCalled -CommandName Test-Path -ParameterFilter { $LiteralPath } -Times 1 -Scope It
            Assert-MockCalled -CommandName Remove-Item -ParameterFilter { $LiteralPath -and $Force } -Times 1 -Scope It
            $err.Exception.Message | Select-Object -First 1 | Should -match 'Failed to convert page'
        }

        It "Publishes OneNote page to Word" {
            Convert-OneNotePage @params 6>$null

            Assert-MockCalled -CommandName Publish-OneNotePage -ParameterFilter { $PublishFormat -eq 'pfWord' } -Times 1 -Scope It
        }

        It "Halts converting if publish OneNote page to Word fails" {
            Mock Publish-OneNotePage -ParameterFilter { $PublishFormat -eq 'pfWord' } { throw }

            $err = Convert-OneNotePage @params 6>$null 2>&1

            $err.Exception.Message | Select-Object -First 1 | Should -match 'Failed to convert page'
        }

        It "Publishes OneNote page to pdf" {
            $params['Config']['exportPdf']['value'] = 2
            Mock Move-Item {}

            Convert-OneNotePage @params 6>$null

            Assert-MockCalled -CommandName Publish-OneNotePage -ParameterFilter { $PublishFormat -eq 'pfPdf' } -Times 1 -Scope It
            Assert-MockCalled -CommandName Move-Item -Times 1 -Scope It
        }

        It "Runs pandoc conversion from docx to markdown" {
            Convert-OneNotePage @params 6>$null

            Assert-MockCalled -CommandName Start-Process -Times 1 -Scope It
        }

        It "Halts converting if pandoc conversion from docx to markdown fails" {
            Mock Start-Process { throw }

            $err = Convert-OneNotePage @params 6>$null 2>&1

            $err.Exception.Message | Select-Object -First 1 | Should -match 'Failed to convert page'
        }

        It "Logs pandoc errors" {
            Mock Start-Process {
                [PSCustomObject]@{
                    ExitCode = 1
                }
            }
            Mock Get-Content {
                'i am some error from pandoc'
            }

            $err = Convert-OneNotePage @params 6>$null 2>&1

            $err.Exception.Message | Select-Object -First 1 | Should -match 'i am some error from pandoc'
        }

        It "Saves page attachment(s)" {
            Convert-OneNotePage @params 6>$null

            Assert-MockCalled -CommandName Copy-Item -Times 1 -Scope It
        }

        It "Does not halt converting if saving of any attachment fails" {
            Mock Copy-Item { throw }

            $err = Convert-OneNotePage @params 6>$null 2>&1

            $err.Exception.Message | Select-Object -First 1 | Should -match 'Error while saving attachment'
        }

        It "Rename page image(s) to unique names" {
            Convert-OneNotePage @params 6>$null

            Assert-MockCalled -CommandName Get-ChildItem -Times 1 -Scope It
            Assert-MockCalled -CommandName Move-Item -Times 1 -Scope It
        }

        It "Does not halt conversion if renaming image(s) fails" {
            Mock Move-Item { throw }

            $err = Convert-OneNotePage @params 6>$null 2>&1

            $err.Exception.Message | Select-Object -First 1 | Should -match 'Error while renaming image'
        }

        It "Markdown Mutation: Rename page image references in markdown to unique names" {
            Convert-OneNotePage @params 6>$null

            Assert-MockCalled -CommandName Get-Content -ParameterFilter { $LiteralPath -and $Raw } -Times 1 -Scope It
            Assert-MockCalled -CommandName Set-ContentNoBom -ParameterFilter { $LiteralPath } -Times 1 -Scope It
        }

        It "Does not halt conversion if renaming image(s) references in markdown fails" {
            Mock Get-Content -ParameterFilter { $LiteralPath -and $Raw } { throw }

            $err = Convert-OneNotePage @params 6>$null 2>&1

            $err.Exception.Message | Select-Object -First 1 | Should -match 'Error while renaming image file name references to'
        }

        It "Markdown mutation: Performs mutations on markdown content" {
            Convert-OneNotePage @params 6>$null

            Assert-MockCalled -CommandName Get-Content -ParameterFilter { $LiteralPath -and ! $Raw } -Times 1 -Scope It
            Assert-MockCalled -CommandName Set-ContentNoBom -ParameterFilter { $LiteralPath } -Times 1 -Scope It
        }

        It "Does a dry run" {
            Mock New-Item { 'foo' }
            Mock Remove-Item { 'foo' }
            Mock Publish-OneNotePage { 'foo' }
            Mock Start-Process { 'foo' }
            Mock Move-Item { 'foo' }
            Mock Get-Content { 'foo' }
            Mock Set-ContentNoBom { 'foo' }

            $params['Config']['DryRun']['value'] = 2

            $result = Convert-OneNotePage @params 6>$null

            $result | Should -Be $null
        }

    }

}

Describe "Print-ConversionErrors" -Tag 'Unit' {

    Context 'Behavior' {

        It "Prints only WriteErrorExceptions of a given array of exceptions" {
            $exceptions = @(
                Write-Error 'foo' 2>&1
                New-Object System.IO.FileNotFoundException -ArgumentList 'bar'
                New-Object System.IO.IOException -ArgumentList 'baz'
            )

            $messages = Print-ConversionErrors -ErrorCollection $exceptions 6>&1

            # Skip the first line
            $messages[1] | Should -Match 'foo'
        }

    }

}

Describe "Convert-OneNote2MarkDown" -Tag 'Unit' {

    Context 'Error Handling' {

        It 'Should throw terminating errors only when -ErrorAction Continue or SilentlyContinue' {
            function Validate-Dependencies {}
            Mock Validate-Dependencies {
                throw
            }

            { Convert-OneNote2MarkDown -ErrorAction Continue 6>$null 2>$null } | Should -Not -Throw
        }

        It 'Should throw terminating errors when -ErrorAction Stop' {
            function Validate-Dependencies {}
            Mock Validate-Dependencies {
                throw
            }

            { Convert-OneNote2MarkDown -ErrorAction Stop 6>$null } | Should -Throw
        }

    }

    Context 'Behavior' {

        BeforeEach {
            $params = @{
                ErrorAction = 'Stop'
            }
            function Validate-Dependencies {}
            Mock Validate-Dependencies {}
            function Compile-Configuration {}
            Mock Compile-Configuration {
                Get-DefaultConfiguration
            }
            function Validate-Configuration {
                param (
                    [Parameter(ValueFromPipeline)]
                    [object]
                    $InputObject
                )
            }
            Mock Validate-Configuration {
                param (
                    [Parameter(ValueFromPipeline)]
                    [object]
                    $InputObject
                )
                process {
                    $InputObject
                }
            }
            function Print-Configuration {}
            Mock Print-Configuration {}
            function New-OneNoteConnection {}
            Mock New-OneNoteConnection { 'some connection' }
            function Get-OneNoteHierarchy {}
            Mock Get-OneNoteHierarchy {
                Get-FakeOneNoteHierarchy
            }
            function New-SectionGroupConversionConfig {}
            Mock New-SectionGroupConversionConfig { 'some conversion config' }
            function Convert-OneNotePage {
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
                    [Parameter(Mandatory,ValueFromPipeline)]
                    [ValidateNotNullOrEmpty()]
                    [object]
                    $InputObject
                )
            }
            Mock Convert-OneNotePage {}
            Mock Get-Variable {}
            function Print-ConversionErrors {}
            Mock Print-ConversionErrors {}
        }

        It "Validates dependencies, compiles and validates configuration, prints configuration, connects to onenote, gets notes hierarchy, builds conversion configuration and converts notes, finally cleans up and prints any errors" {
            Convert-OneNote2MarkDown @params 6>$null

            Assert-MockCalled -CommandName Validate-Dependencies -Times 1 -Scope It
            Assert-MockCalled -CommandName Compile-Configuration -Times 1 -Scope It
            Assert-MockCalled -CommandName Validate-Configuration -Times 1 -Scope It
            Assert-MockCalled -CommandName Print-Configuration -Times 1 -Scope It
            Assert-MockCalled -CommandName New-OneNoteConnection -Times 1 -Scope It
            Assert-MockCalled -CommandName Get-OneNoteHierarchy -Times 1 -Scope It
            Assert-MockCalled -CommandName New-SectionGroupConversionConfig -Times 2 -Scope It
            Assert-MockCalled -CommandName Convert-OneNotePage -Times 2 -Scope It
            Assert-MockCalled -CommandName Get-Variable -Times 1 -Scope It
            Assert-MockCalled -CommandName Print-ConversionErrors -Times 1 -Scope It
        }

        It "Should honor config targetNotebook" {
            Mock Compile-Configuration {
                $config = Get-DefaultConfiguration
                $config['targetNotebook']['value'] = 'test'
                $config
            }

            Convert-OneNote2MarkDown @params 6>$null

            Assert-MockCalled -CommandName New-SectionGroupConversionConfig -Times 1 -Scope It
            Assert-MockCalled -CommandName Convert-OneNotePage -Times 1 -Scope It
        }

    }

}
