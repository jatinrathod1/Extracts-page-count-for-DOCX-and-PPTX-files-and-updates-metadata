# Define SharePoint site and document library details
$siteUrl = "https://futurrizoninterns.sharepoint.com/sites/MentalHealthCareWebApplication1"
$libraryName = "CustomDocumentLibrary"
$libraryDisplayName = "Custom Document Library"
$localPath = "E:\Work FT\Hierarchical_Files_Library_5355_TEST"  

# Connect to SharePoint Online (Interactive login)
Connect-PnPOnline -URL $siteUrl -UseWebLogin

# Check if the document library exists
$library = Get-PnPList -Identity $libraryName -ErrorAction SilentlyContinue
if (-not $library) {
    Write-Host "Creating document library: $libraryDisplayName..."
    New-PnPList -Title $libraryDisplayName -Url $libraryName -Template DocumentLibrary -OnQuickLaunch
} else {
    Write-Host "Document library '$libraryDisplayName' already exists."
}

# Function to get page count for DOCX files
function Get-WordPageCount {
    param ([string]$filePath)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($filePath)
    $pageCount = $doc.ComputeStatistics(2)  # wdStatisticPages = 2
    $doc.Close($false)
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    return $pageCount
}

# Function to get slide count for PPTX files
function Get-PowerPointSlideCount {
    param ([string]$filePath)
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($filePath, $false, $false, $false)
    $slideCount = $presentation.Slides.Count
    $presentation.Close()
    $powerPoint.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
    return $slideCount
}

# Function to create folder structure and upload files to SharePoint Document Library
function Create-SharePointFoldersAndUploadFiles {
    param ([string]$folderPath)

    # Get all subfolders recursively
    $folders = Get-ChildItem -Path $folderPath -Directory -Recurse

    foreach ($folder in $folders) {
        # Get relative path
        $relativePath = $folder.FullName.Replace($localPath, "").TrimStart("\")
        $relativePath = $relativePath -replace "\\", "/"

        Write-Host "Creating folder in SharePoint: $relativePath"

        # Split the relative path into folder levels
        $folderLevels = $relativePath -split "/"
        $currentPath = ""

        foreach ($level in $folderLevels) {
            $parentFolder = $currentPath
            $currentPath = if ($currentPath -eq "") { $level } else { "$currentPath/$level" }

            # Check if folder exists before creating
            $existingFolder = Get-PnPFolder -Url "$libraryName/$currentPath" -ErrorAction SilentlyContinue
            if (-not $existingFolder) {
                Write-Host "Creating folder: $currentPath"

                if ($parentFolder -eq "") {
                    # Create top-level folder inside the document library
                    Add-PnPFolder -Name $level -Folder $libraryName
                } else {
                    # Create subfolder inside the correct parent folder
                    Add-PnPFolder -Name $level -Folder "$libraryName/$parentFolder"
                }
            }
        }

        # Upload PDF, PPT, and Word files in the current folder to SharePoint
        $fileTypes = "*.pdf", "*.pptx", "*.docx"
        foreach ($fileType in $fileTypes) {
            $files = Get-ChildItem -Path $folder.FullName -Filter $fileType
            foreach ($file in $files) {
                $sharePointPath = "$libraryName/$relativePath/$($file.Name)"
                Write-Host "Uploading file: $($file.Name) to $sharePointPath"

                # Upload the file to the corresponding SharePoint folder
                Add-PnPFile -Path $file.FullName -Folder "$libraryName/$relativePath"

                # Get page count for Word and PPT files
                $pageCount = 0
                if ($file.Extension -eq ".docx") {
                    $pageCount = Get-WordPageCount -filePath $file.FullName
                } elseif ($file.Extension -eq ".pptx") {
                    $pageCount = Get-PowerPointSlideCount -filePath $file.FullName
                }

                # Update metadata with page count
                if ($pageCount -gt 0) {
                    $item = Get-PnPListItem -List $libraryName -Fields "FileLeafRef" | Where-Object { $_.FieldValues.FileLeafRef -eq $file.Name } | Select-Object -First 1
                    if ($item) {
                        Set-PnPListItem -List $libraryName -Identity $item.Id -Values @{ "PageCount" = $pageCount }
                        Write-Host "Updated metadata: PageCount = $pageCount for $($file.Name)"
                    } else {
                        Write-Host "Error: Unable to find list item for $($file.Name)"
                    }
                }
            }
        }
    }
}

# Call the function to upload folder structure and files
Write-Host "Creating folder structure and uploading PDF, PPT, and Word files to SharePoint..."
Create-SharePointFoldersAndUploadFiles -folderPath $localPath

Write-Host "Process completed successfully!"
