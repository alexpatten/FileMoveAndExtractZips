<############################################################

This script is used for the following: 
  Moving files uploaded by a vendor via your FTP server and extracting any .ZIP files.
  .ZIP files are copied to an archive folder, and extracted files are moved to a folder for a vendor to access and pull from.
  Logging count of files inside each zip and count of all files moved/extracted.
  Sends email with logs to listed users via SMTP.

############################################################>

[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression')
[Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')

###########################################################
#####           ###########################################
####  VARIABLES  ##########################################
#####           ###########################################
###########################################################

# VARIABLES FOR FILE MOVES FROM FTP TO folder
$tempLogsFolder = "" # Folder to store temporary log files.
$logsFolder = "" # Folder to store all log files.
$initialPath = "" # Initial path of files. 
$destinationPath = "" # Destination path to copy files to.
$archiveDestinationPath = "" # Destination path within archive server.
$zipCounterLogFile = "Zipfile_Count_" # Used for counting zip files from FTP. This name will need to be updated accordingly for each vendor.
# VARIABLES FOR COUNTING FILES IN ZIP FOLDERS
$fileCounterLogFile = "Initial_File_Count_" # Used for counting files in zip folders. This name will need to be updated accordingly for each vendor.
# VARIABLES FOR PROCESSING ZIP FILES
$processingPath = "" # Third path to process zipped files to.
$processingZipsLogFile = "Processed_Zips_" # Log file name
# VARIABLES FOR VALIDATIONS
$validationsLogFile = "Final_Validated_File_Count_" # Log file name
# VARIABLES FOR SENDING EMAIL
$path = '' # Script path
$PasswordFile = $path+'Password.txt' # Variable for encrypted password
$KeyFile = $path+'AES.key' # Variable for encrypted key file
$key = Get-Content $KeyFile # Get content from key file
$user = '' # Variable for email from
$email = @('', '') # Variable for email to
$smtpServer = '' # SMTP server address
# STATIC VARIABLES
$moveCount = 0
$otherFileCount = 0
$zipCounter = 0
$zipFileCount = 0
$procFileCounter = 0
$procZipCounter = 0
$fileCounter = 0
$validationCounter = 0
$badValidationCounter = 0

# If no files exist in FTP folder, close script.
$testPath = Test-Path -Path $initialPath"\*"
if($testPath -eq $False) {
    Write-Host "---- No files in directory. Closing. ----"
    Exit
}

###########################################################
#####           ###########################################
####  FUNCTIONS  ##########################################
#####           ###########################################
###########################################################

Function LogPath($folder, $file) {
    $ts = $(get-date -f yyyyMMdd-HHmmss)
    $newPath = $folder+$file+$ts+"ET.log"
    return $newPath
}

Function CSVCheck {
    $csvList = Get-ChildItem -Path $processingPath -Filter *.csv | ForEach-Object -Process {[System.IO.Path]::GetFileNameWithoutExtension($_)} # Variable for all CSVs in processed folder
    foreach($line in $csvList) { # Loop through CSV list
        $csvPath = $processingPath+"\"+$line+".csv" # Variable for path of CSV in iteration
        $csv = Import-Csv $csvPath 

        [int]$LinesInFile = -1 # Variable for counting rows in CSV
        $reader = New-Object IO.StreamReader $csvPath # Reader object
        while($reader.ReadLine() -ne $null){ $LinesInFile++ } # Count rows
        Write-Host "Number of files in" $line":" $LinesInFile
        
        foreach($row in $csv) { # Loop through each row
            $fileToCheck = $row.Filename # File name of iteration
            if(Test-Path -Path $processingPath"\"$fileToCheck) { # Test the path
                $validationCounter++ # Add to validation counter
                Write-Host "File found: "$processingPath"\"$fileToCheck 
            }
            else { # If path not found
                $badValidationCounter++ # Add to failed counter
                Write-Host "File not found: "$processingPath"\"$fileToCheck 
            }
        }
        $reader.Dispose() # Dispose object to reduce errors
    }

    Write-Host "---- Validation Errors:" $badValidationCounter "files failed validation. ----"
    Write-Host "---- Validation:" $validationCounter "files passed validation. ----"
}

###########################################################
#####            ##########################################
####  MOVE FILES  #########################################
#####            ##########################################
###########################################################

# Starting the log file
$newPath = LogPath $tempLogsFolder $zipCounterLogFile
Start-Transcript -path $newPath -append

# Move files to new folder (This does not change the modified dates)
$files = Get-ChildItem -Path $initialPath -Recurse # Variable for files in FTP folder
ForEach($line in $files) { # Loop through files
    if($line -notmatch ".chk") { # If not CHK file
        if($line -match ".zip") { # If zip file
            $moveCount++ # Add to moveCount
            Write-Host "MOVING ZIP FOLDER #"$moveCount" ----" $line

            Move-Item -Path $initialPath\$line -Destination $destinationPath\$line # Move files to temp_archive folder
            Copy-Item -Path $destinationPath\$line -Destination $archiveDestinationPath\$line # Copy Files to archive server
        }
        else { # If not a zip file
            $otherFileCount++ # Add to otherFileCount
            Write-Host "MOVING UNIQUE FILE #"$otherFileCount" ----" $line

            Move-Item -Path $initialPath\$line -Destination $processingPath\$line # Move files to processed folder
            Copy-Item -Path $processingPath\$line -Destination $archiveDestinationPath\$line # Copy Files to archive server
        }
    }
    else { # If CHK file
        Remove-Item $initialPath\$line # Delete CHK file
    }
}

Write-Host "---- Moved a total of" $moveCount "ZIP folders. ----"
Write-Host "---- Moved a total of" $otherFileCount "unique files. ----"

Stop-Transcript # Stop transcript


###########################################################
#####                     #################################
####  COUNT FILES IN ZIPS  ################################
#####                     #################################
###########################################################

# Zip File Counter
$ErrorActionPreference = "Stop"

# Starting the log file
$newPath = LogPath $tempLogsFolder $fileCounterLogFile
Start-Transcript -path $newPath -append


# working files
$workingfiles = Get-ChildItem -Path $destinationPath\*.zip -Recurse | Select-Object -ExpandProperty FullName
Write-Host $workingfiles

$workingfiles | foreach{
$outfile = $_
$stream = New-Object IO.FileStream($outfile, [IO.FileMode]::Open)
$mode = [IO.Compression.ZipArchiveMode]::Read
$zip = New-Object IO.Compression.ZipArchive($stream, $mode)
$zipCounter ++
Write-Host $(Get-Date) "- Processing $_ - Number of files within:"$zip.Entries.Count "- Zipfile number:"$zipCounter 
$zipFileCount += $zip.Entries.Count
$zip.Dispose()
$stream.Close()
$stream.Dispose()
}

Write-Host "---- Count of all files in each zip:" $zipFileCount "----"

Stop-Transcript


###########################################################
#####                   ###################################
####  PROCESS ZIP FILES  ##################################
#####                   ###################################
###########################################################
$procFileCounter = 0 

# Starting the log file
$newPath = LogPath $tempLogsFolder $processingZipsLogFile
Start-Transcript -path $newPath -append

# Working files
$workingfiles = Get-ChildItem -Path $destinationPath\*.zip -Recurse | Select-Object -ExpandProperty FullName
Write-Host $workingfiles

$workingfiles | foreach {
    $outfile = $_
    $stream = New-Object IO.FileStream($outfile, [IO.FileMode]::Open)
    $mode = [IO.Compression.ZipArchiveMode]::Read
    $zip = New-Object IO.Compression.ZipArchive($stream, $mode)
    $procZipCounter ++
    Write-Host $(Get-Date) "| Processing $_ | Number of files to process:"$zip.Entries.Count "| Zipfile number:"$procZipCounter 

    foreach ($entry in $zip.Entries) {
        $procFileCounter ++
        Write-Host $(Get-Date) "| Extracting $entry TO [$processingPath\$entry] | File #: $procFileCounter" 
        $FileName = $entry.Name
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, "$processingPath\$FileName", $true)
    }

$zip.Dispose()
$stream.Close()
$stream.Dispose()
}

Write-Host "---- Process Complete! - Number of Zip Files Processed:"$zipCounter "- Number of Files Moved:"$procFileCounter "----"

Stop-Transcript


###########################################################
#####                       ###############################
####  COUNT PROCESSED FILES  ##############################
#####                       ###############################
###########################################################

# Count files and store in variable
$processedCount = (Get-ChildItem $processingPath | Measure-Object).Count

# Starting the log file
$newPath = LogPath $tempLogsFolder $validationsLogFile

###########################################################
#####             #########################################
####  VALIDATIONS  ########################################
#####             #########################################
###########################################################

Start-Transcript -path $newPath -append
# Write final results
CSVCheck # Perform validation function
Write-Host "---- The final amount of files that were processed is:" $processedCount "----"
Write-Host "---- The initial amount of counted files is:" $zipFileCount "----"
Write-Host "---- The inital amount of zip folders is:" $moveCount "---- The initial amount of unique files is:" $otherFileCount "----"

Stop-Transcript # Stop transcript

###########################################################
#####                  ####################################
####  SEND LOGS EMAIL   ###################################
#####                  ####################################
###########################################################

$cred = New-Object -TypeName System.Management.Automation.PSCredential ` # Variable for encrypted credentials
    -ArgumentList $user, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)
    
$tempLogList = @() # Array for list of logs attachments

Get-ChildItem -Path $tempLogsFolder -Filter $mask | # Add attachment paths as strings to array
    Foreach-Object {
        $tempLogList += $_.FullName 
    }

$ftemp = @() # Array for attachments
for($i=0; $i -le $tempLogList.length; $i++) {
    $ftemp += $tempLogList[$i]
    }

Send-MailMessage ` # Send email
    -From $user `
    -Subject "Test" `
    -To $email.Split(';') `
    -Attachments 
    -Body "Test message"  `
    -Port 587 `
    -SmtpServer $smtpServer `
    -UseSsl `
    -Credential $cred
    

###########################################################
#####                  ####################################
####  CLEANUP AND EXIT  ###################################
#####                  ####################################
###########################################################

$fileList = Get-ChildItem -Path $destinationPath -Recurse #Variable for files in Temp_Archive folder
# Remove files in temp archive
ForEach($line in $fileList) {
    Remove-Item $destinationPath\$line
}
# Move temp logs to logs folder
ForEach($line in $tempLogList) {
    Move-Item -Path $line -Destination $logsFolder
}

Exit