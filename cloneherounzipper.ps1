Write-Warning ("If there are multiple zip files with the same name, this script will loop forever")
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
$source
$processed = @()
$extractsToNew = 0
$extractsToExisting = 0
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.RootFolder = "MyComputer"
$foldername.Description = "Select a folder for unzipping and extracting"
if($foldername.ShowDialog() -eq "OK") {
    $source = $foldername.SelectedPath
    cd -d $source
}
else {
    Write-Host ("No folder selected. Exiting.")
    Start-Sleep -Seconds 2
    return
}

function CheckDuplicates() {
    if($processed.Contains($zip.BaseName)) {
        Write-Host ("Zip files with same name '" + $zip.BaseName + "' found. Skipping unzipping '" + $zip.FullName +"'")
        return $true
    }
    else{
        return $false
    }
}

#find any .zip files
while($true) {
    $zips = Get-ChildItem -Recurse -Filter "*.zip"
    if($zips.Length -eq 0) {
        break
    }

    #unzip zips
    foreach ($zip in $zips) {
        #test not extracting two zip files with the same name
        if(CheckDuplicates) {
            return
        }
        #Set extraction folder path
        $dest = $zip.FullName.Replace([io.path]::GetExtension($zip), "")
        $test = Test-Path $dest
        if(!$test) {
            Write-Host ("Directory '" + $dest + "' not found. Creating new one.")
            New-Item -ItemType "directory" $dest
            $extractsToNew ++
        }
        else {
            Write-Host ("Existing directory found. Extracting to '" + $dest + "'")
            $extractsToExisting ++
        }
        Expand-Archive $zip.FullName -DestinationPath $dest
        $processed += $zip.BaseName
        Remove-Item $zip.FullName
    }
}

Write-Host ("No new .zip files found. Exiting.")
Write-Host ("New folders created : " + $extractsToNew + "`r`n" +
            "Zip files extracted to existing folders : " + $extractsToExisting)
Remove-Variable source
Remove-Variable processed
Remove-Variable extractsToNew
Remove-Variable extractsToExisting
Remove-Variable foldername
Remove-Variable in
Remove-Variable zips
Remove-Variable zip
Remove-Variable test
Remove-Variable dest
Start-Sleep -Seconds 5
return