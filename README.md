# Convert OneNote to MarkDown

[![github-actions](https://github.com/theohbrothers/ConvertOneNote2MarkDown/actions/workflows/ci-master-pr.yml/badge.svg?branch=master)](https://github.com/theohbrothers/ConvertOneNote2MarkDown/actions/workflows/ci-master-pr.yml)
[![github-release](https://img.shields.io/github/v/release/theohbrothers/ConvertOneNote2MarkDown?style=flat-square)](https://github.com/theohbrothers/ConvertOneNote2MarkDown/releases/)

Ready to make the step to Markdown and saying farewell to your OneNote, EverNote or whatever proprietary note taking tool you are using? Nothing beats clear text, right? Read on!

The powershell script `ConvertOneNote2MarkDown-v2.ps1` will utilize the OneNote Object Model on your workstation to convert all OneNote pages to Word documents and then utilizes Pandoc to convert the Word documents to Markdown (`.md`) format.

## Summary

* Choose between converting a **specific notebook** or **all notebooks** in OneNote. Only notebooks which are [open and synchronized in OneNote](#q-some-notebooks-are-not-detected--converted) can be converted.
* Choose to do a dry run or run the actual conversion.
* Create a **folder structure** for your Notebooks and Sections
  * Process pages that are in sections at the **Notebook, Section Group and all Nested Section Group levels**
* Choose between creating **subfolders for subpages** (e.g. `Page\Subpage.md`) or **appending prefixes** (e.g. `Page_Subpage.md`)
* Specify a value between `32` and `255` as the maximum length of markdown file names, and their folder names (only when using subfolders for subpages (e.g. `Page\Subpage.md`). A lower value can help avoid hitting [file and folder name limits of `255` bytes on file systems](https://en.wikipedia.org/wiki/Comparison_of_file_systems#Limits). A higher value preserves a longer page title. If using page prefixes (e.g. `Page_Subpage.md`), it is recommended to use a value of `100` or greater.
* Choose between putting all media (images, attachments) in a central `/media` folder for each notebook, or in a separate `/media` folder in each folder of the hierarchy
  * Symbols in media file names removed for link compatibility
  * Updates media references in the resulting `.md` files, generating **relative** references to the media files within the markdown document
* Choose between **discarding or keeping intermediate Word files**. Intermediate Word files are stored in a central notebook folder.
* Choose between converting from existing `.docx` (90% faster) and creating new ones - useful if just want to test differences in the various processing options without generating new `.docx`each time
* Choose between naming `.docx` files using page ID and last modified epoch date e.g. `{somelongid}-1234567890.docx` or hierarchy e.g. `<sectiongroup>-<section>-<page>.docx`
* **specify Pandoc output format and any optional extensions**, defaulting to Pandoc Markdown format which strips most HTML from tables and using pipe tables. Customize this to get the desired Markdown flavor. See more details on these options [here](https://pandoc.org/MANUAL.html#options):
  * `markdown` (Pandoc Markdown)
  * `commonmark` (CommonMark Markdown)
  * `gfm` (GitHub-Flavored Markdown), or the deprecated and less accurate markdown_github; use markdown_github only if you need extensions not supported in gfm.
  * `markdown_mmd` (MultiMarkdown)
  * `markdown_phpextra` (PHP Markdown Extra)
  * `markdown_strict` (original unextended Markdown)
  * Improved headers, with title now as a `#` heading, standardized `DateTime` format for created and modified dates, and horizontal line to separate from rest of document
* Choose whether to clear extra newlines between unordered (bullet) and ordered (numbered) list items, empty paragraphs (blank lines containing only non-breaking spaces), non-breaking spaces, and `>` after unordered lists, which are created when converting with Pandoc
* Choose whether to remove `\` escape symbol that are created when converting with Pandoc
* Choose whether to use Line Feed (`LF`) or Carriage Return + Line Feed (`CRLF`) for new lines
* Choose whether to include a `.pdf` export alongside the `.md` file. `.md` does not preserve `InkDrawing` (i.e. overlayed drawings, highlights, pen marks) absolute positions within a page, but a `.pdf` export is a complete page snapshot that preserves `InkDrawing` absolute positions within a page.
* Detailed logs. Run the script with `-Verbose` to see detailed logs of each page's conversion.

## Known Issues

1. If there are any collapsed paragraphs in your pages, the collapsed/hidden paragraphs will not be exported in the final `.md` file
    * You can use the included Onetastic Macro script to automatically expand all paragraphs in each Notebook
    * [Download Onetastic here](https://getonetastic.com/download) and, once installed, use New `Macro > File > Import` to install the attached .xml macro file within Onetastic
1. Password protected sections should be unlocked before continuing, the Object Model does not have access to them if you don't
1. You should start by "flattening" all `InkDrawing` (i.e. pen/hand written elements) in your onennote pages. Because OneNote does not have this function you will have to take screenshots of your pages with pen/hand written notes and paste the resulting image and then remove the scriblings. If you are a heavy pen user this is a very cumbersome.
    * Alternatively, if you are converting a notebook only for reading sake, and want to preserve all notes layout, instead of flattening all `InkDrawing` manually, you may prefer to export a  `.pdf` which preserves the full apperance and layout of the original note (including `InkDrawing`). Simply use the config option `$exportPdf = 2` to export a `.pdf` alongisde the `.md` file.
1. While running the conversion OneNote will be unusable and it is recommended to walk away and have some coffee as the Object Model might be interrupted if you do anything else.
1. Linked file object in `.md` files are clickable in VSCode, but do not open in their associated program, you will have to open the files directly from the file system.

## Requirements

* Windows >= 10

* Windows Powershell 5.x, or [Powershell Core 6.x up to 7.0.x](#q-how-to-install-and-run-powershell-70x). Note:

  * Windows Powershell is already included in Windows 10 / 11 (click Start > `Windows Powershell`). Installing Powershell Core is optional.
  * [Powershell Core 7.1 and above is unsupported](#q-why-is-powershell-core-71-and-above-unsupported).
  * Do not use Windows Powershell ISE, because it [does not support long paths](#error-convert-onenotepage--error-while-renaming-image-file-name-references-to-xxxpng-illegal-characters-in-path).

* Microsoft OneNote >= 2016 (To be clear, this is the Desktop version, not the Windows Store version. Can be downloaded for free [here](https://www.onenote.com/Download)

* Microsoft Word >= 2016 (To be clear, this is the Desktop version NOT the Windows Store version. Can be installed with [Office 365 Trial](https://www.microsoft.com/en-us/microsoft-365/try)

* [Pandoc >= 2.11.2](#q-how-to-install-pandoc)

## Usage

1. Clone or [download](https://github.com/theohbrothers/ConvertOneNote2MarkDown/archive/refs/heads/master.zip) the repository to acquire the powershell script.
1. Start the OneNote application (may require opening as **Adminstrator** in some cases). Keep OneNote open during the conversion.
1. It is advised that you install Onetastic and the attached macro, which will automatically expand any collapsed paragraphs in the notebook. They won't be exported otherwise.
    * To install the macro, click the New Macro Button within the Onetastic Toolbar and then select `File > Import` and select the `.xml` macro included in the release.
    * Run the macro for each Notebook that is open
1. It is highly recommended that you use VS Code, and its embedded Powershell terminal, as this allows you to edit and run the script, as well as check the results of the `.md` output all in one window.
1. If you prefer to use a configuration file, rename `config.example.ps1` to `config.ps1` and configure options in `config.ps1` to your liking.
    * You may like to use `$dryRun = 1` to do a dry run first. This is useful for trying out different settings until you find one you like.
1. Whatever you choose, open Powershell as **Administrator** (or VSCode as **Administrator**, with Powershell as the integrated terminal).

    * In Powershell, first navigate to the folder containing the script:

    ```powershell
    cd "C:\path\to\ConvertOneNote2Markdown"
    ```

    * Then run the script:

    ```powershell
    .\ConvertOneNote2MarkDown-v2.ps1
    ```

    * If you would like to see detailed logs about the conversion process, use the `-Verbose` switch:

    ```powershell
    .\ConvertOneNote2MarkDown-v2.ps1 -Verbose
    ```

    * By default, the script does not terminate on any errors during conversion, but it keeps continuing. If you would like the script to terminate immediately on any error it encounters, use the `-ErrorAction Stop` switch (That way, if the script finishes without errors, you can be sure the entire conversion process went perfectly):

    ```powershell
    .\ConvertOneNote2MarkDown-v2.ps1 -ErrorAction Stop
    ```

    * If you see an [error about scripts being blocked on the system](#error-file-convertonenote2markdownps1-cannot-be-loaded-because-running-scripts-is-disabled-on-this-system), run this line (don't worry, this only allows the current powershell process to bypass security):

    ```powershell
    Set-ExecutionPolicy Bypass -Scope Process -Force
    ```

    * If you see any other errors, check the [FAQ](#faq).
1. If you chose to use a configuration file `config.ps1`, skip to the next step. If you did not choose to use a configuration file, the script will ask you for configuration interactively.
    * It starts off asking whether to do a dry run. This is useful for trying out different settings until you find one you like.
    * It will ask you for the path to store the markdown folder structure. Please use an empty folder. If using VS Code, you might not be able to paste the filepath - right click on the blinking cursor and it will paste from clipboard. Use a full absolute path.
    * Read the prompts carefully to select your desired options. If you aren't actively editing your pages in OneNote, it is highly recommended that you don't delete the intermediate word docs, as they take 80+% of the time to generate. They are stored in their own folder, out of the way. You can then quickly re-run the script with different parameters until you find what you like.
1. Sit back and wait until the process completes.
1. To stop the process at any time, press `Ctrl+C`.
1. If you like, you can inspect some of the `.md` files prior to completion. If you're not happy with the results, stop the process, delete the `.md` and media folders and re-run with different configuration options.

    * If you want to convert to Obsidian Markdown, try any of these recommendations (see [here](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/123), [here](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/150)):

    ```powershell
    $conversion = 'markdown-simple_tables-multiline_tables-grid_tables+pipe_tables-bracketed_spans+native_spans+startnum'
    $conversion = 'markdown_mmd-simple_tables-multiline_tables-grid_tables+pipe_tables-mmd_link_attributes'
    ```

   * If you want to convert to GitHub Flavored Markdown, try this (see [here](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/145)):

    ```powershell
    $conversion = 'gfm+pipe_tables-raw_html'
    ```

   * If you do not want the line of image dimensions after each image, e.g. `{width="12.072916666666666in" height="6.65625in"}` in markdown, try this (see [here](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/145)):

    ```powershell
    $conversion = 'gfm+pipe_tables-raw_html'
    ```

## Results

The script will log any errors encountered during and at the end of its run, so please review, fix and run again if needed.

If you are satisfied check the results with a markdown editor like VSCode. All images should popup just right in the Preview Pane for Markdown files.

## Recommendations

1. I'd like to strongly recommend the [VS Code Foam extension](https://github.com/foambubble/foam-template), which pulls together a selection of markdown-related extensions to become a comprehensive knowledge management tool.
1. I'd also like to recommend [Obsidian.md](http://obsidian.md), which is another fantastic markdown knowledge management tool.
1. Some other VSCode markdown extensions to check out are:

> Note: The bottom three are not really markdown related but are quite obvious.

```powershell
code `
  --install-extension davidanson.vscode-markdownlint `
  --install-extension ms-vscode.powershell-preview `
  --install-extension jebbs.markdown-extended `
  --install-extension telesoho.vscode-markdown-paste-image `
  --install-extension redhat.vscode-yaml `
  --install-extension vscode-icons-team.vscode-icons `
  --install-extension ms-vsts.team
```

## FAQ

### Q: How to install and run Powershell 7.0.x?

A: To install Powershell `7.0.13` (the highest supported version of Powershell) without overridding any existing version of Powershell Core on your system, download [PowerShell-7.0.13-win-x64.zip](https://github.com/PowerShell/PowerShell/releases/download/v7.0.13/PowerShell-7.0.13-win-x64.zip) (validate its checksum [here](https://github.com/PowerShell/PowerShell/releases/v7.0.13)), extract it to a directory `C:\PowerShell-7.0.13-win-x64`, and run `C:\PowerShell-7.0.13-win-x64\pwsh.exe`.

To uninstall after your are done converting, simply delete the `C:\PowerShell-7.0.13-win-x64` directory.

### Q: Why is Powershell Core 7.1 and above unsupported?

A: Powershell `7.1.x` and above no longer supports loading Win32 GAC Assemblies, which is needed for interacting with the OneNote Desktop APIs. It is very unlikely that Microsoft will add support for later Powershell Core versions, since Win32 is starting to become deprecated.

### Q: How to install Pandoc?

A: Pandoc can be installed in any of the following ways:

* Download [pandoc Windows `.msi` installer](https://pandoc.org/installing.html). During installation, ensure **Pandoc is installed for all users** (by checking the box `Install for all users of thie machine`), which is needed for Pandoc to be available to Administrator.
* Use [Chocolatey](https://docs.chocolatey.org/en-us/choco/setup#install-with-powershell.exe) to install the [Pandoc package](https://chocolatey.org/packages/pandoc) on Windows.

After installing, ensure to **restart Powershell or the computer** for Pandoc to be set correctly in `PATH`.

### Q: Error(s) when opening OneNote as Administrator

A: If there are errors opening OneNote as Administrator, just open it normally without Administrator permissions. A user has reported successful conversion with only Powershell opened as Administrator. See [case](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/149#issuecomment-1418051753).

### Q: Some notebooks are not detected / converted

A: The script cannot detect notebooks which are not yet open in OneNote. Use `File > Open` in OneNote to open all OneNote notebooks that you want to convert, and ensure they are all fully synchronized before running the conversion. This applies to local OneNote notebooks and cloud OneNote notebooks (E.g. OneDrive or Microsoft Teams).

### Q: Page attachments (files, images) are missing or corrupt during conversion

Cause: Possible page content bug.

A: Create a new section, copy pages into it, run the script again. See [case](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/183), which is similar to this [case](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/112#issuecomment-986947168).

### Error: `File ConvertOneNote2Markdown.ps1 cannot be loaded because running scripts is disabled on this system.`

A: Windows disables Powershell (`.ps1`) scripts by default for security reasons. To allow running `.ps1` scripts temporarily only for the current Powershell session, run:

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
```

### Error: `Error HRESULT E_FAIL has been returned from a call to a COM component`

Cause: Powershell `7.1.x` and above does not support loading Win32 GAC Assemblies.

Solution: Use a version of Powershell between `5.x` and `7.0.x`. See [here](#q-how-to-install-and-run-powershell-70x).

### Error: `80080005 Server execution failed (Exception from HRESULT: 0x80080005(CO_E_SERVER_EXEC_FAILURE)`

Cause: Mismatch in security contexts of Powershell and OneNote.

Solution: Ensure both Powershell and OneNote are run under the same user privileges. An easy way is to run both Powershell and OneNote as Administrator.

### Error: `Unable to find type [Microsoft.Office.InterOp.OneNote.HierarchyScope]`

Cause: Mismatch in security contexts of Powershell and OneNote.

Solution: Ensure both Powershell and OneNote are run under the same user privileges. An easy way is to run both Powershell and OneNote as Administrator.

### Error: `Exception calling "Publish" with "4" argument(s): "Class not registered"`

Cause: Microsoft Word is not installed.

Solution: Ensure Microsoft Word is installed.

### Error: `Exception: Exception calling "Publish" with "4" argument(s): "Server execution failed  Server execution failed"`

Cause: Microsoft Word plugin(s) causing publishing to `.docx` to fail. E.g. `zotero` plugin (see [case](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/149#issuecomment-1418051753)).

Solution: Disable all Microsoft Word plugins before running the conversion.

### Error: `Exception: Exception calling "Publish" with "4" argument(s): "0x80042006"`

Cause 1: Mismatch in security contexts of Powershell and OneNote.

Solution 1: Ensure both Powershell and OneNote are run under the same user privileges. An easy way is to run both Powershell and OneNote as Administrator.

Cause 2: `$notesdestpath` is not an absolute path.

Solution 2: Use an absolute path for `$notesdestpath`.

### Error: `Exception calling "Publish" with "4" argument(s): "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)"`

Cause: Windows Defender could be blocking `pandoc.exe`. See [case](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/11#issuecomment-835081893).

Solution: Disable Windows Defender. Click `Start > Windows Security`, under `Virus and threat protection > Exclusions` click `Add or remove exclusions`, click `Add an exclusion` and browse to `C:\Program Files\Pandoc\pandoc.exe`. Then retry the conversion.

If `pandoc.exe` is still blocked, you might try to turn off settings under `Virus and threat protection`, `App and browser control`, and `Device Security`.

### Error: `Exception calling "Publish" with "4" argument(s): "The remote procedure call failed. (Exception from HRESULT: 0x800706BE)"`

Cause 1: OneNote is not open during the conversion.

Solution 1: Open OneNote and keep it open during the conversion.

Cause 2: : Page content bug.

Solution 2: Create a new section, copy pages into it, run the script again. See [case](https://github.com/theohbrothers/ConvertOneNote2MarkDown/issues/112#issuecomment-986947168).

### Error: `Convert-OneNotePage : Error while renaming image file name references to 'xxx.png: Illegal characters in path.`

Cause: Windows Powershell ISE does not support long paths using the [`\\?\`](https://docs.microsoft.com/en-us/windows/win32/fileio/naming-a-file?redirectedfrom=MSDN#maxpath) prefix, e.g. `\\?\C:\path\to`.

Solution: Do not use Window Powershell ISE. Use Windows Powershell, or Powershell Core. See [requirements](#requirements).

## Credits

* Avi Aryan for the awesome [VSCodeNotebook](https://github.com/aviaryan/VSCodeNotebook) port
* [@SjoerdV](https://github.com/SjoerdV) for the [original script](https://github.com/SjoerdV/ConvertOneNote2MarkDown)
* [@nixsee](https://github.com/nixsee) who made a variety of modifications and improvements on the fork, which was transferred to [@theohbrothers](https://github.com/theohbrothers)

<!--
title:  'Convert OneNote to MarkDown'
author:
- Sjoerd de Valk, SPdeValk Consultancy
- modified by nixsee, a guy
- modified by The Oh Brothers
date: 2020-07-15 13:00:00
keywords: [migration, tooling, onenote, markdown, powershell]
abstract: |
  This document is about converting your OneNote data to Markdown format.
permalink: /index.html
-->
