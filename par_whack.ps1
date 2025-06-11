# Set the root directory to start the search
$rootDirectory = "\\smb-101.storage.ausps.org\des_library$\Library\"

# Create an array to hold the results for CSV export
$results = @()

# Get all *.*.par files recursively in the directory
Get-ChildItem -Path $rootDirectory -Recurse -Filter "*.*.par" | ForEach-Object {
    # Generate the new filename by removing the .par extension
    $newName = $_.FullName -replace '\.par$', ''

    # Rename the file
    Rename-Item -Path $_.FullName -NewName $newName

    # Add the old and new filenames to the results array
    $results += [PSCustomObject]@{
        old_filename = $_.FullName
        new_filename = $newName
    }

    # Output the renamed files to the console
    Write-Host "Renamed: $($_.FullName) to $newName"
	
	catch {
        # Handle any errors and display the error message
        Write-Host "Error renaming file: $($_.FullName)"
        Write-Host "Error Message: $($_.Exception.Message)"
        
        # Pause the script to allow the user to see the error
        Read-Host -Prompt "Press Enter to continue after the error"
    }
}

# Check if the $results array contains any data
if ($results.Count -gt 0) {
    # Export the results array to a CSV file
    $csvPath = "$env:USERPROFILE\Desktop\par_clean.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation

    Write-Host "CSV export completed: $csvPath"
} else {
    Write-Host "No files were renamed. No CSV was generated."
}