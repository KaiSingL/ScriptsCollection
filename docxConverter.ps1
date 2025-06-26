# Requires Microsoft Word to be installed
function Cleanup-Word {
    param ($word)
    try {
        if ($null -ne $word) {
            $word.Quit()
        }
    }
    catch {
        Write-Host "Warning: Error closing Word application: $_"
    }
    finally {
        if ($null -ne $word) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    while ($true) {
        # Prompt for directory path
        $directory = Read-Host 'Please provide the directory path in double quotes (or press Enter to exit)'
        
        # Exit if no path provided
        if ([string]::IsNullOrWhiteSpace($directory)) {
            Write-Host "Exiting script"
            break
        }
        
        # Remove quotes from the input path
        $directory = $directory.Trim('"')
        
        # Validate directory exists
        if (-not (Test-Path $directory)) {
            Write-Host "Error: Directory '$directory' does not exist"
            continue
        }
        
        # Get all .doc files in the specified directory
        $docFiles = Get-ChildItem -Path $directory -Filter "*.doc" -File
        
        if ($docFiles.Count -eq 0) {
            Write-Host "No .doc files found in the specified directory"
            continue
        }
        
        foreach ($docFile in $docFiles) {
            $document = $null
            try {
                # Open the document
                $document = $word.Documents.Open($docFile.FullName)
                
                # Remove all macros if they exist
                if ($document.VBProject.VBComponents.Count -gt 0) {
                    foreach ($component in $document.VBProject.VBComponents) {
                        $document.VBProject.VBComponents.Remove($component)
                    }
                }
                
                # Create new file name with .docx extension
                $newFileName = [System.IO.Path]::ChangeExtension($docFile.FullName, ".docx")
                
                # Save as docx (FileFormat 16 is wdFormatXMLDocument)
                $document.SaveAs([ref]$newFileName, [ref]16)
                
                Write-Host "Converted: $($docFile.Name) -> $(Split-Path $newFileName -Leaf)"
            }
            catch {
                Write-Host "Error processing $($docFile.Name): $_"
            }
            finally {
                if ($null -ne $document) {
                    try {
                        $document.Close()
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
                    }
                    catch {
                        Write-Host "Warning: Error closing document $($docFile.Name): $_"
                    }
                }
            }
        }
        
        Write-Host "Conversion complete for directory: $directory"
    }
}
catch {
    Write-Host "Critical error: $_"
}
finally {
    Cleanup-Word -word $word
    Remove-Variable word -ErrorAction SilentlyContinue
}