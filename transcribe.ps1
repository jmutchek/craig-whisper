# A PowerShell script to automate the transcription of audio files.
# Uses OpenAI's whisper implementation and Python
#
# Parameters
#  - inputFolder    - folder that contains all of the audio files to be transcribed
#                   - filenames follow a naming convention `n-playername_m` where n and m are numbers
#  - outputFolder   - folder for all transcription output, defaults to `<inputFolder>\..\transcriptions`
#  - force          - switch parameter to force re-transcription of already processed files
#  - postProcessOnly - switch parameter to skip transcription and only perform post-processing
#  - cleanup        - switch parameter to remove temporary files after processing; if not provided, temp files are kept
#  - ignoreWords    - array of words/phrases to be filtered out from transcriptions (case-insensitive)
# 
# Requirements
#  - intended to be run unattended on a Windows machine with Python and whisper installed
#  - Python must be in the PATH environment variable
#  - whisper must be in the PATH environment variable (install via `pip install -U openai-whisper`)
#  - ffmpeg must be in the PATH environment variable
#  - the audio files must be in a format supported by ffmpeg (e.g. mp3, wav, m4a, flac, ogg, aac, mp4, wma)
#  - the expected naming convention for audio files is `n-playername_m.ext` where n and m are numbers
#    - if files don't match this pattern, the filename without extension will be used as the speaker name
#  - the script uses the large-v2 model by default with English language processing
#  - the script will continue to the next file if there is an error in transcription
#  - all errors are logged in the transcription_log.csv and transcription_state.csv files
#  - the script tracks transcription progress in a state file to resume interrupted transcriptions
#  - statistics are collected about duration, file sizes, and processing time
#  - output format:
#    - individual TSV files for each transcription with speaker, start time, end time, and text
#    - original whisper files are stored in the archive directory
#    - processed files are stored in the temp directory
#  - the -Force parameter can be used to override CSV column mismatches when appending errors
#  - the -PostProcessOnly parameter can be used to skip transcription and only run the post-processing
#  - the -Cleanup parameter can be used to remove temporary files after processing
#  - the -IgnoreWords parameter can be used to specify words to filter out from transcriptions
#
# Usage examples:
#  - .\transcribe.ps1 -InputFolder "C:\path\to\audio\files"
#  - .\transcribe.ps1 -InputFolder "C:\path\to\audio\files" -OutputFolder "C:\path\to\output"
#  - .\transcribe.ps1 -InputFolder "C:\path\to\audio\files" -Force
#  - .\transcribe.ps1 -InputFolder "C:\path\to\audio\files" -PostProcessOnly
#  - .\transcribe.ps1 -InputFolder "C:\path\to\audio\files" -Cleanup
#  - .\transcribe.ps1 -InputFolder "C:\path\to\audio\files" -IgnoreWords @("you", "um", "uh")
#
# Examples
#  - This is an example of the transcription output file:
#    ```
#   speaker  start   end     text
#   jmutchek 26480   32960   I'll be able to do much easier sort of speaker categorization after the fact for the transcript,
#   jmutchek 32960   39920   see what happens. Okay, but it's just audio, it's not going to record anything else.
#   jmutchek 41840   52960   What about, okay, I told you I'm going to sort of, I'm 100% stealing slash riffing on Betsy's
#   jmutchek 52960   63600   backstory, and I know she is on a journey to find her, the owner of her secret library.
#    ```
# 
# Process
#  - for each audio file
#    - transcribe using the whisper command `whisper --model large-v2 --language en --condition_on_previous_text False --compression_ratio_threshold 1.8 --output_dir <OUTPUT_DIR> <AUDIO_FILE>`
#    - add a tab-separated column to the start of each line in the transcription file with the playername extracted from the audio filename
#  - if -Cleanup parameter is specified, temporary files are removed after processing

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true, HelpMessage="Folder containing audio files to transcribe")]
    [string]$InputFolder,
    
    [Parameter(Mandatory=$false, HelpMessage="Output folder for transcriptions")]
    [string]$OutputFolder,
    
    [Parameter(Mandatory=$false, HelpMessage="Force re-transcription of already processed files")]
    [switch]$Force = $false,
    
    [Parameter(Mandatory=$false, HelpMessage="Skip transcription and only perform post-processing on existing transcripts")]
    [switch]$PostProcessOnly = $false,
    
    [Parameter(Mandatory=$false, HelpMessage="Clean up temporary files after processing")]
    [switch]$Cleanup = $false,
    
    [Parameter(Mandatory=$false, HelpMessage="Words or phrases to be filtered out from transcriptions (case-insensitive)")]
    [string[]]$IgnoreWords = @("you", "silence", "um", "uh", "ah", "like", "right", "well")
)

#region Functions
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp,$Level,$Message"
    Add-Content -Path $logFilePath -Value $logEntry -Encoding UTF8
    
    # Also write to console with color
    switch ($Level) {
        "INFO" { Write-Host $Message -ForegroundColor White }
        "WARNING" { Write-Host "WARNING: $Message" -ForegroundColor Yellow }
        "ERROR" { Write-Host "ERROR: $Message" -ForegroundColor Red }
    }
}

function Test-Dependencies {
    $dependencies = @("python", "whisper", "ffmpeg")
    $missingDeps = @()
    
    foreach ($dep in $dependencies) {
        if (-not (Get-Command $dep -ErrorAction SilentlyContinue)) {
            $missingDeps += $dep
        }
    }
    
    if ($missingDeps.Count -gt 0) {
        Write-Log "Missing dependencies: $($missingDeps -join ', ')" "ERROR"
        return $false
    }
    
    return $true
}

function Get-PlayerName {
    param (
        [string]$FileName
    )
    
    if ($FileName -match "^\d+-([^_]+)_\d+") {
        return $Matches[1]
    }
    
    # If file doesn't match our pattern, use filename as player name
    return [System.IO.Path]::GetFileNameWithoutExtension($FileName)
}

function Transcribe-AudioFile {
    param (
        [string]$AudioFile,
        [string]$TempOutputDir
    )
    
    $fileName = [System.IO.Path]::GetFileName($AudioFile)
    $fileSize = (Get-Item $AudioFile).Length
    $playerName = Get-PlayerName $fileName
    $startTime = Get-Date
    
    Write-Log "Starting transcription of $fileName (Size: $([math]::Round($fileSize/1MB, 2)) MB)" "INFO"
    
    try {
        # Call whisper for transcription
        $whisperArgs = "--model large-v2 --language en --condition_on_previous_text False --compression_ratio_threshold 1.8 --output_dir `"$TempOutputDir`" `"$AudioFile`""
        $process = Start-Process -FilePath "whisper" -ArgumentList $whisperArgs -NoNewWindow -PassThru -Wait
        
        if ($process.ExitCode -ne 0) {
            throw "Whisper process exited with code $($process.ExitCode)"
        }
        
        # Get the original whisper output file (should be a .tsv file)
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($AudioFile)
        $originalTsvFile = Join-Path $TempOutputDir "$baseFileName.tsv"
        
        if (-not (Test-Path $originalTsvFile)) {
            throw "Expected transcript file not found: $originalTsvFile"
        }
        
        # Create a copy in the archive directory to preserve original whisper output
        $originalWhisperOutput = Join-Path $archiveDir "$baseFileName.whisper_original.tsv"
        Copy-Item -Path $originalTsvFile -Destination $originalWhisperOutput
        
        Write-Log "Original Whisper output archived at: $originalWhisperOutput" "INFO"
        
        # Add the player name to each line in the transcription
        $transcriptContent = Get-Content $originalTsvFile -Encoding UTF8
        $updatedContent = @()
        
        foreach ($line in $transcriptContent) {
            if ($line -match "^start") {
                # Update header line to include speaker column
                $updatedContent += "speaker`t$line"
            }
            elseif ([string]::IsNullOrWhiteSpace($line)) {
                # Keep empty lines unchanged
                $updatedContent += $line
            }
            else {
                # Add playername as first column
                $updatedContent += "$playerName`t$line"
            }
        }
        
        # Save processed transcript to temp directory only
        $processedFile = Join-Path $TempOutputDir "$baseFileName.processed.tsv"
        $updatedContent | Out-File -FilePath $processedFile -Encoding UTF8
        
        # Calculate processing time and stats
        $endTime = Get-Date
        $duration = ($endTime - $startTime).TotalSeconds
        
        # Record successful transcription in state file
        $stateEntry = [PSCustomObject]@{
            FileName = $fileName
            FileSize = $fileSize
            ProcessingTime = [math]::Round($duration, 2)
            Status = "Success"
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            PlayerName = $playerName
            ErrorMessage = ""  # Add empty ErrorMessage to maintain consistent schema
        }
        
        $stateEntry | Export-Csv -Path $stateFilePath -Append -NoTypeInformation -Encoding UTF8 -Force
        
        Write-Log "Completed transcription of $fileName in $([math]::Round($duration, 2)) seconds" "INFO"
        return $true
    }
    catch {
        # Record failed transcription in state file
        $endTime = Get-Date
        $duration = ($endTime - $startTime).TotalSeconds
        
        $stateEntry = [PSCustomObject]@{
            FileName = $fileName
            FileSize = $fileSize
            ProcessingTime = [math]::Round($duration, 2)
            Status = "Error"
            ErrorMessage = $_.Exception.Message
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            PlayerName = $playerName
        }
        
        $stateEntry | Export-Csv -Path $stateFilePath -Append -NoTypeInformation -Encoding UTF8 -Force
        
        Write-Log "Error transcribing ${fileName}: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Process-WhisperFile {
    param (
        [string]$WhisperFile,
        [string]$TempOutputDir,
        [string]$OutputFolder
    )
    
    $fileName = [System.IO.Path]::GetFileName($WhisperFile)
    $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    
    # Handle different file naming patterns based on where the file is coming from
    # If it's already an archive file with .whisper_original.tsv extension, use the base name without that suffix
    if ($fileName -match "\.whisper_original\.tsv$") {
        $originalBaseFileName = $baseFileName -replace "\.whisper_original$", ""
        $playerName = Get-PlayerName -FileName $originalBaseFileName
    } else {
        $playerName = Get-PlayerName -FileName $baseFileName
    }
    
    Write-Log "Processing whisper file: $fileName for player: $playerName" "INFO"
    
    try {
        # Read the original whisper output file
        $transcriptContent = Get-Content -Path $WhisperFile -Encoding UTF8
        
        # Create a copy in the archive directory to preserve original whisper output, but only if it doesn't already exist
        $originalWhisperOutput = Join-Path $archiveDir "$baseFileName.whisper_original.tsv"
        
        # If this file is already from the archive (has .whisper_original.tsv extension), 
        # or if it's already in the archive, don't create another copy
        if ($fileName -match "\.whisper_original\.tsv$") {
            # File is already an archive file, use its path as the originalWhisperOutput
            $originalWhisperOutput = $WhisperFile
            Write-Log "Using existing archive file: $fileName" "INFO"
        } 
        elseif (-not (Test-Path $originalWhisperOutput)) {
            # Only copy if the file doesn't already exist in the archive
            Copy-Item -Path $WhisperFile -Destination $originalWhisperOutput
            Write-Log "Original Whisper output archived at: $originalWhisperOutput" "INFO"
        }
        else {
            Write-Log "Archive file already exists, skipping copy: $originalWhisperOutput" "INFO"
        }
        
        $processedLines = New-Object System.Collections.ArrayList
        
        # Add header line
        $headerLine = ($transcriptContent | Where-Object { $_ -match "^start" } | Select-Object -First 1)
        if ($headerLine) {
            [void]$processedLines.Add("speaker`t$headerLine")
        } else {
            [void]$processedLines.Add("speaker`tstart`tend`ttext") # Default header if none found
        }
        
        # Process content lines and store as objects for easier manipulation
        $contentLines = @()
        $filteredCount = 0
        
        foreach ($line in $transcriptContent) {
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }
            
            if ($line -match "^start") {
                # Skip header line as we've already added it
                continue
            }
            
            # Extract the text portion and other fields
            $lineParts = $line -split '\t'
            
            # Skip lines with fewer than expected columns
            if ($lineParts.Count -lt 3) {
                continue
            }
            
            $startTime = $lineParts[0]
            $endTime = $lineParts[1]
            $text = $lineParts[2].Trim()
            
            # Check if the text matches any of the ignore words/phrases (with or without punctuation)
            $shouldFilter = $false
            foreach ($ignoreWord in $IgnoreWords) {
                # Check for exact match or match with trailing punctuation
                if ($text -ieq $ignoreWord -or $text -imatch "^$([regex]::Escape($ignoreWord))[.,!?;:]?$") {
                    $shouldFilter = $true
                    $filteredCount++
                    break
                }
            }
            
            if (-not $shouldFilter) {
                $contentLines += [PSCustomObject]@{
                    Speaker = $playerName
                    Start = $startTime
                    End = $endTime
                    Text = $text
                }
            }
        }
        
        # Collapse consecutive identical lines
        $collapsedLines = New-Object System.Collections.ArrayList
        $collapsedCount = 0
        
        if ($contentLines.Count -gt 0) {
            # Initialize with the first item
            $currentGroup = [PSCustomObject]@{
                Speaker = $contentLines[0].Speaker
                Start = $contentLines[0].Start
                End = $contentLines[0].End
                Text = $contentLines[0].Text
            }
            
            # Process the remaining items starting from the second item
            for ($i = 1; $i -lt $contentLines.Count; $i++) {
                $current = $contentLines[$i]
                
                # Check if the current line has the same speaker and text as the current group (case-insensitive comparison)
                if ($current.Speaker -eq $currentGroup.Speaker -and $current.Text -ieq $currentGroup.Text) {
                    # Update the end timestamp of the current group to the latest end time
                    $currentGroup.End = $current.End
                    $collapsedCount++
                } else {
                    # Add the current group to our collapsed results
                    [void]$collapsedLines.Add($currentGroup)
                    
                    # Start a new group with the current item
                    $currentGroup = [PSCustomObject]@{
                        Speaker = $current.Speaker
                        Start = $current.Start
                        End = $current.End
                        Text = $current.Text
                    }
                }
            }
            
            # Add the last group
            [void]$collapsedLines.Add($currentGroup)
        }
        
        # Convert collapsed lines back to TSV format and add to processed lines
        foreach ($line in $collapsedLines) {
            [void]$processedLines.Add("$($line.Speaker)`t$($line.Start)`t$($line.End)`t$($line.Text)")
        }
        
        # Save processed transcript directly to the temp directory only
        $processedFile = Join-Path $TempOutputDir "$baseFileName.processed.tsv"
        $processedLines | Out-File -FilePath $processedFile -Encoding UTF8
        
        Write-Log "Created processed file: $processedFile (processed $($contentLines.Count) lines, collapsed $collapsedCount, filtered $filteredCount)" "INFO"
        return $true
    }
    catch {
        Write-Log "Error processing whisper file ${fileName}: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-ProcessingStats {
    param (
        [string]$StateFilePath
    )
    
    if (-not (Test-Path $StateFilePath)) {
        return @{
            TotalFiles = 0
            SuccessCount = 0
            ErrorCount = 0
            TotalDuration = 0
            AverageDuration = 0
            TotalDurationFormatted = "0m 0s"
            AverageDurationFormatted = "0m 0s"
            TotalSize = 0
        }
    }
    
    $stateData = Import-Csv -Path $StateFilePath -Encoding UTF8
    
    $totalSeconds = ($stateData | Measure-Object -Property ProcessingTime -Sum).Sum
    
    $stats = @{
        TotalFiles = $stateData.Count
        SuccessCount = ($stateData | Where-Object { $_.Status -eq "Success" }).Count
        ErrorCount = ($stateData | Where-Object { $_.Status -eq "Error" }).Count
        TotalDuration = [math]::Round($totalSeconds, 2)
        AverageDuration = 0
        TotalDurationFormatted = ""
        AverageDurationFormatted = ""
        TotalSize = [math]::Round(($stateData | Measure-Object -Property FileSize -Sum).Sum / 1MB, 2)
    }
    
    # Format total duration as minutes and seconds
    $totalMinutes = [math]::Floor($totalSeconds / 60)
    $remainingSeconds = [math]::Round($totalSeconds % 60, 0)
    $stats.TotalDurationFormatted = "${totalMinutes}m ${remainingSeconds}s"
    
    if ($stats.TotalFiles -gt 0) {
        $avgSeconds = $totalSeconds / $stats.TotalFiles
        $stats.AverageDuration = [math]::Round($avgSeconds, 2)
        
        # Format average duration as minutes and seconds
        $avgMinutes = [math]::Floor($avgSeconds / 60)
        $avgRemainingSeconds = [math]::Round($avgSeconds % 60, 0)
        $stats.AverageDurationFormatted = "${avgMinutes}m ${avgRemainingSeconds}s"
    } else {
        $stats.AverageDurationFormatted = "0m 0s"
    }
    
    return $stats
}

function Create-AggregateTranscripts {
    param (
        [string]$TempOutputDir,
        [string]$OutputFolder
    )
    
    Write-Log "Starting post-post-processing: Creating aggregate transcript files" "INFO"
    
    # Check if there are any processed transcript files
    $processedFiles = Get-ChildItem -Path $TempOutputDir -Filter "*.processed.tsv" -ErrorAction SilentlyContinue
    
    if ($processedFiles.Count -gt 0) {
        Write-Log "Found $($processedFiles.Count) processed transcript files to aggregate" "INFO"
        
        try {
            # Create a merged file sorted by start time
            $mergedTranscriptFile = Join-Path $OutputFolder "merged_transcript.tsv"
            $mergedLines = New-Object System.Collections.ArrayList
            $allContentLines = @()
            
            # Add header only once
            [void]$mergedLines.Add("speaker`tstart`tend`ttext")
            
            # Read all processed files and combine their contents
            foreach ($file in $processedFiles) {
                $content = Get-Content -Path $file.FullName -Encoding UTF8
                
                # Skip the header line of each file
                $contentWithoutHeader = $content | Where-Object {
                    -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch "^speaker\t"
                }
                
                if ($contentWithoutHeader.Count -gt 0) {
                    $allContentLines += $contentWithoutHeader
                }
            }
            
            # Convert lines to objects for sorting
            $transcriptEntries = @()
            
            foreach ($line in $allContentLines) {
                $parts = $line -split "`t"
                
                # Skip lines with fewer than 4 parts (speaker, start, end, text)
                if ($parts.Count -lt 4) {
                    continue
                }
                
                # Convert start time to numeric for proper sorting
                $startTime = [int]::Parse($parts[1])
                
                $transcriptEntries += [PSCustomObject]@{
                    Line = $line
                    StartTime = $startTime
                    Speaker = $parts[0]
                }
            }
            
            # Sort by start time
            $sortedEntries = $transcriptEntries | Sort-Object -Property StartTime
            
            # Add sorted lines to the merged file
            foreach ($entry in $sortedEntries) {
                [void]$mergedLines.Add($entry.Line)
            }
            
            # Write the merged file to the main output directory - this is the primary output
            $mergedLines | Out-File -FilePath $mergedTranscriptFile -Encoding UTF8
            
            $transcriptCount = $sortedEntries.Count
            Write-Log "Created merged transcript file with $transcriptCount entries: $mergedTranscriptFile" "INFO"
            
            # Create speaker-specific aggregate files in the temp directory
            $speakers = $sortedEntries | Select-Object -ExpandProperty Speaker -Unique
            
            foreach ($speaker in $speakers) {
                # Write speaker files to temp directory instead of main output
                $speakerFile = Join-Path $TempOutputDir "$speaker`_transcript.tsv"
                $speakerLines = New-Object System.Collections.ArrayList
                
                # Add header
                [void]$speakerLines.Add("speaker`tstart`tend`ttext")
                
                # Get lines for this speaker
                $speakerEntries = $sortedEntries | Where-Object { $_.Speaker -eq $speaker }
                
                # Add lines to speaker file
                foreach ($entry in $speakerEntries) {
                    [void]$speakerLines.Add($entry.Line)
                }
                
                # Write the speaker file to the temp directory
                $speakerLines | Out-File -FilePath $speakerFile -Encoding UTF8
                
                $entryCount = $speakerEntries.Count
                Write-Log "Created speaker transcript file for $speaker with $entryCount entries in temp directory: $speakerFile" "INFO"
            }
            
            return $true
        }
        catch {
            Write-Log "Error creating aggregate transcript files: $($_.Exception.Message)" "ERROR"
            return $false
        }
    } else {
        Write-Log "No processed transcript files found to aggregate" "WARNING"
        return $false
    }
}
#endregion

#region Main Script
# Set default output folder if not specified
if (-not $OutputFolder) {
    $OutputFolder = Join-Path (Split-Path $InputFolder -Parent) "transcriptions"
}

# Check if input folder exists
if (-not (Test-Path $InputFolder)) {
    Write-Host "Input folder does not exist: $InputFolder" -ForegroundColor Red
    exit 1
}

# Create output folder if it doesn't exist
if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
    Write-Host "Created output folder: $OutputFolder"
}

# Set up logging and state tracking
$logFilePath = Join-Path $OutputFolder "transcription_log.csv"
$stateFilePath = Join-Path $OutputFolder "transcription_state.csv"
$tempOutputDir = Join-Path $OutputFolder "temp"
$archiveDir = Join-Path $OutputFolder "archive"

# Create temp and archive directories if they don't exist
if (-not (Test-Path $tempOutputDir)) {
    New-Item -Path $tempOutputDir -ItemType Directory | Out-Null
}

if (-not (Test-Path $archiveDir)) {
    New-Item -Path $archiveDir -ItemType Directory | Out-Null
}

# Initialize log file if it doesn't exist
if (-not (Test-Path $logFilePath)) {
    "Timestamp,Level,Message" | Out-File -FilePath $logFilePath -Encoding UTF8
}

# Initialize state file if it doesn't exist
if (-not (Test-Path $stateFilePath)) {
    "FileName,FileSize,ProcessingTime,Status,Timestamp,PlayerName,ErrorMessage" | Out-File -FilePath $stateFilePath -Encoding UTF8
}

# Log script start
Write-Log "Starting transcription process" "INFO"
Write-Log "Input folder: $InputFolder" "INFO"
Write-Log "Output folder: $OutputFolder" "INFO"
Write-Log "Force re-transcription: $Force" "INFO"
Write-Log "Post-process only: $PostProcessOnly" "INFO"
Write-Log "Cleanup temp folder: $Cleanup" "INFO"
Write-Log "Words to ignore: $($IgnoreWords -join ', ')" "INFO"

# Check for dependencies
if (-not $PostProcessOnly) {
    if (-not (Test-Dependencies)) {
        Write-Log "Missing required dependencies. Please ensure Python, Whisper, and FFmpeg are installed and in PATH." "ERROR"
        exit 1
    }
}

if (-not $PostProcessOnly) {
    # Get audio files
    $audioFiles = Get-ChildItem -Path $InputFolder -File | Where-Object {
        $_.Extension -match "\.(mp3|wav|m4a|flac|ogg|aac|mp4|wma)$"
    }

    if ($audioFiles.Count -eq 0) {
        Write-Log "No supported audio files found in $InputFolder" "WARNING"
        exit 0
    }

    Write-Log "Found $($audioFiles.Count) audio files to process" "INFO"

    # Get already processed files from state file
    $processedFiles = @()
    if (Test-Path $stateFilePath) {
        $stateData = Import-Csv -Path $stateFilePath -Encoding UTF8
        $processedFiles = $stateData | Where-Object { $_.Status -eq "Success" } | Select-Object -ExpandProperty FileName
    }

    $successCount = 0
    $errorCount = 0

    # Process each audio file
    foreach ($audioFile in $audioFiles) {
        if (-not $Force -and $processedFiles -contains $audioFile.Name) {
            Write-Log "Skipping already processed file: $($audioFile.Name)" "INFO"
            continue
        }
        
        $result = Transcribe-AudioFile -AudioFile $audioFile.FullName -TempOutputDir $tempOutputDir
        
        if ($result) {
            $successCount++
        } else {
            $errorCount++
        }
    }
} else {
    Write-Log "Running in post-process only mode - skipping transcription" "INFO"
    
    # Clean up any existing processed files first to avoid duplication
    $existingProcessedFiles = Get-ChildItem -Path $tempOutputDir -Filter "*.processed.tsv" -ErrorAction SilentlyContinue
    if ($existingProcessedFiles.Count -gt 0) {
        Write-Log "Removing $($existingProcessedFiles.Count) existing processed files before recreating them" "INFO"
        $existingProcessedFiles | Remove-Item -Force
    }
    
    # Check if there are any whisper transcript files in BOTH the temp AND archive directories
    $whisperFilesTmp = Get-ChildItem -Path $tempOutputDir -Filter "*.tsv" -ErrorAction SilentlyContinue | 
                      Where-Object { 
                          $_.Name -notmatch "\.processed\.tsv$" -and 
                          $_.Name -notmatch "\.whisper_original\.tsv$" 
                      }
    
    $whisperFilesArchive = Get-ChildItem -Path $archiveDir -Filter "*.whisper_original.tsv" -ErrorAction SilentlyContinue
    
    $whisperFiles = @()
    
    # Add files from temp directory if any
    if ($whisperFilesTmp.Count -gt 0) {
        $whisperFiles += $whisperFilesTmp
    }
    
    # Add files from archive directory if any
    if ($whisperFilesArchive.Count -gt 0) {
        $whisperFiles += $whisperFilesArchive
    }
    
    if ($whisperFiles.Count -eq 0) {
        Write-Log "No Whisper transcription files found in temp or archive directories" "WARNING"
        Write-Host "No Whisper transcription files found to post-process in: $tempOutputDir or $archiveDir" -ForegroundColor Yellow
        exit 0
    }
    
    Write-Log "Found $($whisperFiles.Count) Whisper transcription files to process" "INFO"
    
    # Recreate .processed.tsv files from whisper output
    $successCount = 0
    $errorCount = 0
    
    foreach ($file in $whisperFiles) {
        $result = Process-WhisperFile -WhisperFile $file.FullName -TempOutputDir $tempOutputDir -OutputFolder $OutputFolder
        
        if ($result) {
            $successCount++
        } else {
            $errorCount++
        }
    }
    
    # Check if any processed files were created
    $transcriptFiles = Get-ChildItem -Path $tempOutputDir -Filter "*.processed.tsv" -ErrorAction SilentlyContinue
    
    if ($transcriptFiles.Count -eq 0) {
        Write-Log "No processed transcript files were created in temp directory: $tempOutputDir" "WARNING"
        Write-Host "Failed to create any processed transcript files in: $tempOutputDir" -ForegroundColor Yellow
        exit 0
    }
    
    Write-Log "Successfully created $($transcriptFiles.Count) processed transcript files" "INFO"
}

# Calculate and display stats
$stats = Get-ProcessingStats -StateFilePath $stateFilePath

# Generate summary report
Write-Log "Transcription process complete" "INFO"
Write-Log "Total files processed: $($stats.TotalFiles)" "INFO"
Write-Log "Successfully transcribed: $($stats.SuccessCount)" "INFO"
Write-Log "Failed transcriptions: $($stats.ErrorCount)" "INFO"
Write-Log "Total processing time: $($stats.TotalDurationFormatted)" "INFO"
Write-Log "Average processing time per file: $($stats.AverageDurationFormatted)" "INFO"
Write-Log "Total audio size processed: $($stats.TotalSize) MB" "INFO"

# Create aggregate transcript files
Write-Log "Starting aggregate transcript file creation" "INFO"
$aggregateResult = Create-AggregateTranscripts -TempOutputDir $tempOutputDir -OutputFolder $OutputFolder

if ($aggregateResult) {
    Write-Log "Successfully created aggregate transcript files in output folder" "INFO"
} else {
    Write-Log "No aggregate transcript files were created" "WARNING"
}

Write-Host ""
Write-Host "Transcription Summary:" -ForegroundColor Cyan
Write-Host "======================" -ForegroundColor Cyan
Write-Host "Total files processed: $($stats.TotalFiles)"
Write-Host "Successfully transcribed: $($stats.SuccessCount)" -ForegroundColor Green
if ($PostProcessOnly) {
    Write-Host "Successfully post-processed: $successCount" -ForegroundColor Green
    Write-Host "Failed post-processing: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
} else {
    Write-Host "Failed transcriptions: $($stats.ErrorCount)" -ForegroundColor $(if ($stats.ErrorCount -gt 0) { "Red" } else { "Green" })
}
Write-Host "Total processing time: $($stats.TotalDurationFormatted)"
Write-Host "Average time per file: $($stats.AverageDurationFormatted)"
Write-Host "Total audio size: $($stats.TotalSize) MB"
Write-Host "Output folder: $OutputFolder"
Write-Host "Individual .processed.tsv files have been created in the output folder" -ForegroundColor Cyan

# Display aggregate file information
if ($aggregateResult) {
    Write-Host ""
    Write-Host "Aggregate Output Files:" -ForegroundColor Cyan
    Write-Host "----------------------" -ForegroundColor Cyan
    Write-Host "* merged_transcript.tsv - All transcripts sorted by start time" -ForegroundColor Green
    
    # Look for speaker-specific files in temp directory now instead of output folder
    $speakerFiles = Get-ChildItem -Path $TempOutputDir -Filter "*_transcript.tsv" -ErrorAction SilentlyContinue
    
    if ($speakerFiles.Count -gt 0) {
        Write-Host "Speaker-specific transcript files in temp directory:" -ForegroundColor Cyan
        foreach ($file in $speakerFiles) {
            $speaker = [System.IO.Path]::GetFileNameWithoutExtension($file.Name).Replace("_transcript", "")
            Write-Host "* $($file.Name) - Speaker-specific transcript for $speaker" -ForegroundColor Green
        }
    }
}

Write-Host ""

# Clean up temp directory
if ($Cleanup) {
    Write-Log "Cleaning up temporary directory: $tempOutputDir" "INFO"
    Remove-Item -Path $tempOutputDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Temporary files cleaned up" -ForegroundColor Cyan
} else {
    Write-Log "Temporary directory kept (use -Cleanup to remove): $tempOutputDir" "INFO"
    Write-Host "Temporary files kept at: $tempOutputDir" -ForegroundColor Cyan
}
#endregion

