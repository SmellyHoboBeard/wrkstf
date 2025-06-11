# Set the root directory to start the search
$rootDirectory = "\\smb-101.storage.ausps.org\des_library$\Library\"
$csvPath = "$env:USERPROFILE\Desktop\par_clean.csv"

# Get all *.*.par files, rename them, and collect the results.
# The -PassThru parameter on Rename-Item outputs the renamed file object.
# The results of the loop are directly assigned to the $renamedFiles variable.
$renamedFiles = Get-ChildItem -Path $rootDirectory -Recurse -Filter "*.*.par" | ForEach-Object {
    $originalPath = $_.FullName
    # Use -replace on the base name of the file for a cleaner new name
    $newName = $_.Name -replace '\.par$', ''

    try {
        # Rename the item and pass the resulting object down the pipeline
        Rename-Item -Path $originalPath -NewName $newName -ErrorAction Stop -PassThru

        # Output progress to the console
        Write-Host "SUCCESS: Renamed '$originalPath' to '$newName'"
    }
    catch {
        # Handle any errors for the specific file and display the message
        Write-Warning "ERROR: Failed to rename '$originalPath'. Message: $($_.Exception.Message)"
    }
}

# Check if any files were successfully renamed before creating the CSV
if ($renamedFiles) {
    # Create a custom report from the renamed file objects
    $report = $renamedFiles | Select-Object @{Name='old_filename'; Expression={$_.FullName.Replace($_.Name, "$($_.Name).par")}}, @{Name='new_filename'; Expression={$_.FullName}}

    # Export the report to a CSV file
    $report | Export-Csv -Path $csvPath -NoTypeInformation

    Write-Host "`nCSV export completed: $csvPath"
} else {
    Write-Host "`nNo files were renamed or found. No CSV was generated."
}

# Optional: Pause at the end
Read-Host -Prompt "Script finished. Press Enter to exit"