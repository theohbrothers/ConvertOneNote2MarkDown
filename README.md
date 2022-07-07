# Convert OneNote to MarkDown

[![github-actions](https://github.com/theohbrothers/ConvertOneNote2MarkDown/workflows/ci-master-pr/badge.svg)](https://github.com/theohbrothers/ConvertOneNote2MarkDown/actions)
[![github-release](https://img.shields.io/github/v/release/theohbrothers/ConvertOneNote2MarkDown?style=flat-square)](https://github.com/theohbrothers/ConvertOneNote2MarkDown/releases/)

Ready to make the step to Markdown and saying farewell to your OneNote, EverNote or whatever proprietary note taking tool you are using? Nothing beats clear text, right? Read on!

The powershell script `ConvertOneNote2MarkDown-v2.ps1` will utilize the OneNote Object Model on your workstation to convert all OneNote pages to Word documents and then utilizes PanDoc to convert the Word documents to Markdown (.md) format.

## Summary

* Choose to do a dry run or run the actual conversion.
* Create a **folder structure** for your Notebooks and Sections
  * Process pages that are in sections at the **Notebook, Section Group and all Nested Section Group levels**
* Choose between converting a **specific notebook** or **all notebooks**
* Choose between creating **subfolders for subpages** (e.g. `Page\Subpage.md`) or **appending prefixes** (e.g. `Page_Subpage.md`)
* Specify a value between `32` and `255` as the maximum length of markdown file names, and their folder names (only when using subfolders for subpages (e.g. `Page\Subpage.md`)). A lower value can help avoid hitting [file and folder name limits of `255` bytes on file systems](https://en.wikipedia.org/wiki/Comparison_of_file_systems#Limits). A higher value preserves a longer page title. If using page prefixes (e.g. `Page_Subpage.md`), it is recommended to use a value of `100` or greater.
* Choose between putting all media (images, attachments) in a central `/media` folder for each notebook, or in a separate `/media` folder in each folder of the hierarchy
  * Symbols in media file names removed for link compatibility
  * Updates media references in the resulting `.md` files, generating **relative** references to the media files within the markdown document
* Choose between **discarding or keeping intermediate Word files**. Intermediate Word files are stored in a central notebook folder.
* Choose between converting from existing `.docx` (90% faster) and creating new ones - useful if just want to test differences in the various processing options without generating new `.docx`each time
* Choose between naming `.docx` files using page ID and last modified epoch date e.g. `{somelongid}-1234567890.docx` or hierarchy e.g. `<sectiongroup>-<section>-<page>.docx`
* **specify Pandoc output format and any optional extensions**, defaulting to Pandoc Markdown format which strips most HTML from tables and using pipe tables. See more details on these options here: https://pandoc.org/MANUAL.html#options
   * markdown (Pandoc’s Markdown)
   * commonmark (CommonMark Markdown)
   * gfm (GitHub-Flavored Markdown), or the deprecated and less accurate markdown_github; use markdown_github only if you need extensions not supported in gfm.
   * markdown_mmd (MultiMarkdown)
   * markdown_phpextra (PHP Markdown Extra)
   * markdown_strict (original unextended Markdown)
* Choose whether to include page timestamp and a separator at top of page
  * Improved headers, with title now as a # heading, standardized DateTime format for created and modified dates, and horizontal line to separate from rest of document
* Choose whether to remove double spaces between bullet points, non-breaking spaces from blank lines, and `>` after bullet lists, which are created when converting with Pandoc
* Choose whether to remove `\` escape symbol that are created when converting with Pandoc
* Choose whether to use Line Feed (LF) or Carriage Return + Line Feed (CRLF) for new lines
* Choose whether to include a `.pdf` export alongside the `.md` file. `.md` does not preserve `InkDrawing` (i.e. overlayed drawings, highlights, pen marks) absolute positions within a page, but a `.pdf` export is a complete page snapshot that preserves `InkDrawing` absolute positions within a page.
* Detailed logs. Run the script with `-Verbose` to see detailed logs of each page's conversion.

## Known Issues

1. If there are any collapsed paragraphs in your pages, the collapsed/hidden paragraphs will not be exported in the final `.md` file
    * You can use the included Onetastic Macro script to automatically expand all paragraphs in each Notebook
    * [Download Onetastic here](https://getonetastic.com/download) and, once installed, use New Macro-> File-> Import to install the attached .xml macro file within Onetastic
1. Password protected sections should be unlocked before continuing, the Object Model does not have access to them if you don't
1. You should start by 'flattening' all `InkDrawing` (i.e. pen/hand written elements) in your onennote pages. Because OneNote does not have this function you will have to take screenshots of your pages with pen/hand written notes and paste the resulting image and then remove the scriblings. If you are a heavy 'pen' user this is a very cumbersome.
    - Alternatively, if you are converting a notebook only for reading sake, and want to preserve all notes layout, instead of flattening all `InkDrawing` manually, you may prefer to export a  `.pdf` which preserves the full apperance and layout of the original note (including `InkDrawing`). Simply use the config option `$exportPdf = 2` to export a `.pdf` alongisde the `.md` file.
1. While running the conversion OneNote will be unusable and it is recommended to 'walk away' and have some coffee as the Object Model might be interrupted if you do anything else.
1. Linked file object in `.md` files are clickable in VSCode, but do not open in their associated program, you will have to open the files directly from the file system.

## Requirements

* Windows >= 10

* Windows Powershell 5.x and above, or [Powershell Core 6.x up to 7.0](https://github.com/PowerShell/PowerShell)

* Microsoft OneNote >= 2016 (To be clear, this is the Desktop version NOT the Windows Store version. Can be downloaded for FREE here - https://www.onenote.com/Download)

* Microsoft Word >= 2016 (To be clear, this is the Desktop version NOT the Windows Store version. Can be installed with Office 365 Trial - https://www.microsoft.com/en-us/microsoft-365/try).

* [PanDoc >= 2.7.2](https://pandoc.org/installing.html)

  * TIP: You may also use [Chocolatey](https://chocolatey.org/docs/installation#install-with-powershellexe) to install Pandoc on Windows, this will also set the right path (environment) statements. (https://chocolatey.org/packages/pandoc)

## Usage

1. Clone this repository to acquire the powershell script.
1. Start the OneNote application
1. It is advised that you install Onetastic and the attached macro, which will automatically expand any collapsed paragraphs in the notebook. They won't be exported otherwise.
    * To install the macro, click the New Macro Button within the Onetastic Toolbar and then select File -> Import and select the .xml macro included in the release.
    * Run the macro for each Notebook that is open
1. It is highly recommended that you use VS Code, and its embedded Powershell terminal, as this allows you to edit and run the script, as well as check the results of the .md output all in one window.
1. If you prefer to use a configuration file, rename `config.example.ps1` to `config.ps1` and configure options in `config.ps1` to your liking.
   1. You may like to use `$dryRun = 2` to do a dry run first. This is useful for trying out different settings until you find one you like.
1. Whatever you choose, open a PowerShell terminal and navigate to the folder containing the script and run it.
    ```.\ConvertOneNote2MarkDown-v2.ps1```
    * If you would like to see detailed logs about the conversion process, use the `-Verbose` switch:
    ```.\ConvertOneNote2MarkDown-v2.ps1 -Verbose```
    * If you have trouble, try running both Onenote and Powershell as an administrator.
    * If you receive an error, try running this line to bypass security:
    ``Set-ExecutionPolicy Bypass -Scope Process``
1. If you chose to use a configuration file `config.ps1`, skip to the next step. If you did not choose to use a configuration file, the script will ask you for configuration interactively.
    * It starts off asking whether to do a dry run. This is useful for trying out different settings until you find one you like.
    * It will ask you for the path to store the markdown folder structure. Please use an empty folder. If using VS Code, you might not be able to paste the filepath - right click on the blinking cursor and it will paste from clipboard. Use a full absolute path.
    *  Read the prompts carefully to select your desired options. If you aren't actively editing your pages in Onenote, it is HIGHLY recommended that you don't delete the intermediate word docs, as they take 80+% of the time to generate. They are stored in their own folder, out of the way. You can then quickly re-run the script with different parameters until you find what you like.
1. Sit back and wait until the process completes
1. To stop the process at any time, press Ctrl+C.
1. If you like, you can inspect some of the .md files prior to completion. If you're not happy with the results, stop the process, delete the .md and media folders and re-run with different parameters.

## Results

The script will log any errors encountered during and at the end of its run, so please review, fix and run again if needed.
If you are satisfied check the results with a markdown editor like VSCode. All images should popup just right in the Preview Pane for Markdown files.

## Recommendations
1. I'd like to strongly recommend the [VS Code Foam extension](https://github.com/foambubble/foam-template), which pulls together a selection of markdown-related extensions to become a comprehensive knowledge management tool.
1. I'd also like to recommend [Obsidian.md](http://obsidian.md), which is another fantastic markdown knowledge management tool.
1. Some other VSCode markdown extensions to check out are:

```powershell
    .\code `
    --install-extension davidanson.vscode-markdownlint `
    --install-extension ms-vscode.powershell-preview `
    --install-extension jebbs.markdown-extended `
    --install-extension telesoho.vscode-markdown-paste-image `
    --install-extension redhat.vscode-yaml `
    --install-extension vscode-icons-team.vscode-icons `
    --install-extension ms-vsts.team
```

> NOTE: The bottom three are not really markdown related but are quite obvious.

## Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

### [2.3] - 2020-07-16
#### Added
* Improved file header with title as # heading, standardized DateTime format and horizontal line to separate from document
* Option to re-use existing docx file - 90% faster - for when testing different processing options

#### Changed
* Tweaks to filenames and other formatting

#### Removed
* Nothing

### [2.2] - 2020-07-16
#### Added
* Symbols removed from attached file names
* Files stored in same folder as images
* Double spaces and "\" escape symbols removed after conversion with Pandoc
* Tables use piped formatting for compatibility with Obsidian

#### Changed
* User prompt layouts

#### Removed
* Nothing

### [2.1] - 2020-07-15
#### Added
* Prompt for keep or discard .docx files
* Prompt to have images in central folder or separate ones for each folder in hierarchy

#### Changed
* User prompt layouts

#### Removed
* Nothing


### [2.0] - 2020-07-14
#### Added
* Consolidated prior scripts into one
* Prompt for markdown format selection
* Prompt to choose between prefix and subfolders for subpages

#### Changed
* Now produces relative references to images (e.g ../../media
* Each notebook has a centralized images/media folder

#### Removed
* Extraneous code

### [1.1] - 2020-07-11
#### Added
 * Two new scripts to allow for Section Groups, as well as Section Groups + Subfolders for Subpages

#### Changed
* Pandoc instead of gfm set as default format

#### Removed
* Nothing

### [1.0.0] - 2019-05-19

#### Added

* Initial Release

#### Changed

* Nothing

#### Removed

* Nothing


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
