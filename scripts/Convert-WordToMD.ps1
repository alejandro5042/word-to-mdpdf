<#
.SYNOPSIS
    Converts a Word document to Markdown and PDF.

.DESCRIPTION
    This script converts a Word document to a Jekyll Markdown site and a PDF.
    It uses the pandoc and wkhtmltopdf utilities to do the conversion.

    The script will create a new directory in the current directory with the
    same name as the Word document. It will then create a Markdown file for
    each 2nd-level (##) header in the Word document.
#>
param (
    # The path to the Word document to convert to Markdown.
    $Path = "originals/vpp-4.3.6.docx"
)

function Convert-ImageFormat {
    <#
    .SYNOPSIS
        Converts an image to a different format.
    .REMARKS
        This is a wrapper around System.Drawing.Bitmap.Save() method.
    #>
	param (
        # The path to the image to convert.
        [Parameter(Mandatory = $true)]
        [string] $Path,

        # The new extension to use for the image.
        [string] $NewExtension = "png"
    )

    $file = Get-Item $Path

    $null = [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $bitmap = new-object System.Drawing.Bitmap($file.Fullname)
    $newFileName = "$($file.DirectoryName)/$($file.BaseName).$NewExtension"
    $bitmap.Save($newFileName, $NewExtension)
}

try {
    $inputFile = Get-Item $Path
    $name = $inputFile.BaseName

    # Recreate the output directory and cd into it.
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue out/$name
    $null = mkdir out -Force
    $outPath = mkdir out/$name -Force
    Push-Location $outPath

    # Convert the Word document to Markdown.
    pandoc ($inputFile.FullName) -t gfm --table-of-contents --wrap=none -o raw.md --extract-media "."

    # Convert images from whatever weird format they are originally in to PNG.
    foreach ($image in Get-ChildItem -Recurse -Filter *.wmf) {
        Convert-ImageFormat $image
    }

    # Fix image paths in the Markdown file.
    foreach ($line in Get-Content "raw.md") {
        $line = $line -replace '\./media/', 'media/'
        $line = $line -replace '\.wmf', '.png'
        $line | Out-File -Append -Encoding UTF8 "$name.md"
    }

    # From now on, we'll be working with the single-page.md file.
    Remove-Item "raw.md"

    # Split the single-page.md into multiple files, starting with the index.md file.
    $path = "index.md"
    Get-Content "$PSScriptRoot/default-metadata.yml" | Out-File -Append -Encoding UTF8 $path
    $chapter = 0
    foreach ($line in Get-Content "$name.md") {
        # If the line is a 2nd-level (##) header, create a new file for it.
        if ($line -match "^##\s+(.*)\s*$") {
            $header = $Matches[1]
            $headerPath = "$outPath/$($chapter.ToString("000"))-$($header -replace "[^a-z0-9]", "-").md"
            $chapter++
            $path = $headerPath
            Get-Content "$PSScriptRoot/default-metadata.yml" | Out-File -Append -Encoding UTF8 $path
        }

        # Write the line to the current file.
        $line | Out-File -Append -Encoding UTF8 $path
    }

    # Copy the base Jekyll site files into the output directory.
    Copy-Item -Recurse -Force "$PSScriptRoot/../base_site/*" .

    # Create a PDF from all the Markdown files, with page numbers.
    # More options on wkhtmltopdf: https://wkhtmltopdf.org/usage/wkhtmltopdf.txt
    pandoc -i "$name.md" -o "$name.pdf" `
        --metadata "title=$name" `
        --css="$PSScriptRoot/pdf.css" `
        --number-sections `
        --pdf-engine=wkhtmltopdf `
        --pdf-engine-opt="--header-left" --pdf-engine-opt="[title]" `
        --pdf-engine-opt="--header-right" --pdf-engine-opt="[section]" `
        --pdf-engine-opt="--footer-center" --pdf-engine-opt="[page] of [topage]" `
        --pdf-engine-opt="--header-spacing" --pdf-engine-opt="10" `
        --pdf-engine-opt="--footer-spacing" --pdf-engine-opt="10" `
        --pdf-engine-opt="--header-font-size" --pdf-engine-opt="10" `
        --pdf-engine-opt="--footer-font-size" --pdf-engine-opt="10"

    # Add links to the bottom of the index.md file.
    "`n`n- **[View as Single Page]($name.md)**" | Out-File -Append -Encoding UTF8 "index.md"
    "`n`n- **[View as PDF]($name.pdf)**" | Out-File -Append -Encoding UTF8 "index.md"

    # Tell the user how to run the site.
    Write-Host "`nTo run the site, run the following command from the root of the repository:"
    Write-Host "    ./scripts/serve.ps1 `"$outPath`""
} catch {
    Write-Error $_
} finally {
    Pop-Location
}
