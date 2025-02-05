# Load the necessary assemblies
Add-Type -AssemblyName System.Windows.Forms

# Display welcome message
Write-Host " `
--------------------------------------------------------------------------------- `
|\/\/\/\|      Hello! This script will crawl a public site that         |/\/\/\/| ` 
|\/\/\/\|      contains download links and download files where the     |/\/\/\/| ` 
|\/\/\/\|      link text matches a user-provided list of entries.       |/\/\/\/| ` 
---------------------------------------------------------------------------------`n"

# Set global varialbes 
$firstRun = $true

# Prompt for url
function Prompt-For-Url {
	$url = $Host.UI.Prompt("Source URL", "Please enter the URL where the download links are displayed.", "URL")
	
	# Make connection, test status, and get html
	try {
		$request = [System.Net.WebRequest]::Create($url["URL"])
		$response = $request.GetResponse()
		
		# Read the response stream
		$stream = $response.GetResponseStream()
		$reader = New-Object System.IO.StreamReader($stream)
		$htmlContent = $reader.ReadToEnd()
	}
	catch {
		Write-Error $_.Exception.Message
		Prompt-For-Url
	}
	finally {
		# Close request
		If ($response -ne $null) { $response.Close() }
	}
	return @($htmlContent, $url)
}

# Prompt for destination folder
function Prompt-For-Folder {
    # Create and configure the folder browser dialog
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "Please select a folder where files will be downloaded to."
    $folderBrowserDialog.ShowNewFolderButton = $true

    # Show the folder browser dialog and get the result
    $dialogResult = $folderBrowserDialog.ShowDialog()

    # Check if the user selected a folder
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedFolder = $folderBrowserDialog.SelectedPath
        return $selectedFolder
    } else {
        Write-Host "No folder was selected."
        exit
    }
}

# Prompt user for if they want to check their list against an existing folder for matches first
function Prompt-For-Cross-Check {
	$title    = "Cross-Check Folder Location?"
	$question = 'Would you like to provide a folder location to cross-check your list against to ensure you do not already have the items downloaded?'

	$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
	$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
	$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

	$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
	if ($decision -eq 0) {
		$message = "  = User cross-check requested."
		Add-Content -Path $logFilePath -Value $message
		return $true
	} else {
		Write-Host ""
		Add-Content -Path $logFilePath -Value "  = No cross check"
		return $false
	}
}

# Clean an array of values as desired
function Clean-Array {
	param (
		[array]$inputArray
	)
	
	#clean array, removing any special characters and spaces.
	$cleanArray = @()
	foreach ($item in $inputArray) {
		$newItem = $item.ToLower() -replace "\(.*$", "" -replace "(\s?(T|t)he | of | or | is | a | an )", "" -replace "[^A-Za-z0-9]", "" | Where-Object {$_}
		$cleanArray += $newItem
	}
	return $cleanArray
}

# Clean a single item of an array as desired
function Clean-Item {
	param (
		[string]$Item
	)
	return $item.ToLower() -replace "\(.*$", "" -replace "(\s?(T|t)he | of | or | is | a | an )", "" -replace "[^A-Za-z0-9]", "" | Where-Object {$_}
}

# Extract the links from the html string
function Extract-Links {
	# Define a regular expression pattern to match the entire anchor (<a>) tag
	$pattern = '<a[^>]*>(.*?)<\/a>'
	$matches = [regex]::Matches($html, $pattern)

	# Create custom objects for each match
	$links = $matches | ForEach-Object {
		$hrefPattern = 'href="([^"]+)"'
		$href = [regex]::Match($_.Value, $hrefPattern).Groups[1].Value
		$linkText = $_.Groups[1].Value
		[PSCustomObject]@{
			Href = $href
			Text = $linkText
		}
	}
	return $links
}

# Get input file, populate array, delete file
function Process-Input {
	Write-Host -ForegroundColor White "`nInput File Contents"
	$ErrorActionPreference = "Stop"
	try {
		$path = "./input/input.txt"
		$inputFileContent = Get-Content -Path $path
		$arrayList = $inputFileContent -split [Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries
		Remove-Item $path
		$message = "Input found!"
		Write-Host $message
		Add-Content -Path $logFilePath -Value $message
	}
	catch {
		Write-Error -Message "Error: $($_.Exception.Message)"
		exit
	}
	return $arrayList
}

# Create the output file to store array of items that did not match
function Create-Output-File {
    $folderPath = "./input"
    if (-not (Test-Path $folderPath)) {
        # Create the folder structure
        New-Item -ItemType Directory -Path $folderPath -Force
    }
    $filePath = $folderPath + "/input.txt"
    
    # Remove empty items from the arrayList
    $arrayList = $arrayList | Where-Object { ![string]::IsNullOrWhiteSpace($_) }
    
    # Convert the arrayList to a single string with newline characters between items
    $outputContent = $arrayList -join "`n"
    
    # Write the cleaned arrayList to the output file without extra lines
    Set-Content -Path $filePath -Value $outputContent -NoNewLine
    Write-Host "  ~ Input File updated."
	Add-Content -Path $logFilePath -Value "Input file updated with remaining values."
}

# Iterate throuh links and download as appropriate
function Download-Links {
	Write-Host "Crawling provided URL..."
	foreach($link in $links) {
		Write-Host "`r $($links.IndexOf($link))" -NoNewLine
		#$baseUrl.URL = $baseUrl.URL.Trim("/")
		$url = "$($baseUrl.URL)$($link.Href)"
		$removeValue = $false
		# only process links that have (USA) in the title and do not contain Beta or Demo
		if ($link.Text -like "*(USA)*" -and $link.Text -notmatch "(\(|\[)(Beta.*|Demo.*)(\)|\[)") {
			Add-Content -Path $logFilePath -Value "- Checking $($link.CleanText) ($($link.Text))"
			#only download when a match occurs
			if ($cleanInputArray -contains $link.CleanText) {
				$message = "`n  * Matched $($link.Text)"
				Write-Host $message
				Add-Content -Path $logFilePath -Value $message
				#I/O setup
				$fileName = [System.IO.Path]::GetFileName($link.Text)
				$destinationPath = Join-Path -Path $destinationFolder -ChildPath $fileName
				$outFile = "$destinationFolder\$fileName"
				
				if (Test-Path $outFile) {
					$message = "  ! File exists: $fileName"
					Write-Host $message
					Add-Content -Path $logFilePath -Value $message
					$removeValue = $true
				}
				else {
					try {
						$message = "  ... Downloading $($link.Text) : $((Invoke-WebRequest $url -Method Head).Headers.'Content-Length') bytes ..."
						Write-Host $message
						Add-Content -Path $logFilePath -Value $message
						
						# Bits Transfer is faster than Invoke WebRequest and allows for a progress bar
						Start-BitsTransfer -Source $url -Destination $outFile

						$message = "  + Downloaded $outFile."
						Write-Host $message
						Add-Content -Path $logFilePath -Value $message
						$removeValue = $true
					}
					catch {
						$errorMessage = "An error occurred: $($_.Exception.Message)`n"
						Write-Error -Message $errorMessage
						Add-Content -Path $logFilePath -Value $errorMessage
					}
				}
			}
		}
		# remove value from the array if the file already exists or was successfully downloaded
		if ($removeValue) { 
			$arrayList[$cleanInputArray.IndexOf($link.CleanText)] = ""
			Create-Output-File
			Write-Host "Crawling provided URL..."
		}
	}
	Write-Host "... Finished crawling URL."
}	

function Prompt-For-More-URLs {
	$title    = "Another URL?"
	$question = 'Would you like to crawl another URL using the same input list and download location?'

	$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
	$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
	$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

	$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
	if ($decision -eq 0) {
		Main
	} 	
}

function Main {
	# Initialize log file
	$logFilePath = ".\log.txt"
	$start = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$initMessage = " `
	--------------------------------------------------------------------------------- `
	|\/\/\/\|              BATCH START at $start         	|/\/\/\/| ` 
	---------------------------------------------------------------------------------`n"
	if (-Not (Test-Path $logFilePath)) {
		# Create the file and write the log message
		Out-File -FilePath $logFilePath -InputObject $initMessage -encoding utf8
		
	} else {
		# Append the log message to the existing file
		Add-Content -Path $logFilePath -Value $initMessage
	}

	# Prompt for url, verify status, and assign html based on url and parse links
	$html, $baseUrl = $(Prompt-For-Url)
	$logHeader = "baseUrl: $baseUrl"
	if (-not $firstRun) { $logHeader += " -- continued..." }

	# Prompt for destination folder
	if ($firstRun) {
		Write-Host -ForegroundColor White "`nDestination Folder"
		$destinationFolder = $(Prompt-For-Folder)
		Write-Host "Destination path: $destinationFolder"
		$logHeader += "`nDestination Folder: $destinationFolder`n"
		Add-Content -Path $logFilePath -Value $logHeader     	# Append base information to the log
	}

	# Clean the link text and add it to the custom object
	$links = $(Extract-Links)
	foreach ($link in $links) {
		$link | Add-Member -MemberType NoteProperty -Name "CleanText" -Value $(Clean-Item($link.Text))
	}

	# Clean the input array
	$arrayList = $(Process-Input)
	$cleanInputArray = Clean-Array($arrayList)

	# Prompt for cross-Check and process accordingly
	if ($firstRun) {
		if (Prompt-For-Cross-Check) {
			$title    = "Same as Download Folder?"
			$question = 'Would you like to use the same folder as where the items are downloaded to?'

			$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
			$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
			$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

			$decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
			if ($decision -eq 0) {
				$crossCheckFolder = $destinationFolder
			} else {
				$crossCheckFolder = $(Prompt-For-Folder)
			}
			
			Write-Host "`nChecking folder for already downloaded items..."
			$files = (Get-ChildItem -Path $crossCheckFolder -File).Name
			$cleanFileNames = $(Clean-Array($files))
			foreach ($file in $cleanFileNames) {
				if ($cleanInputArray -contains $file) {
					$message = "  * Found $($arrayList[$cleanInputArray.IndexOf($file)]).`n  ~ Removing entry from input list.`n"
					Add-Content -Path $logFilePath -Value $message
					Write-Host $message
					$arrayList[$cleanInputArray.IndexOf($file)] = ""
					$cleanInputArray[$cleanInputArray.IndexOf($file)] = ""
				}
			}
		}
	}

	# Download links where they match the input file
	
	Download-Links

	# Save unmatched items to an output file	
	$arrayList = $arrayList | Where-Object { $_ } | Where-Object { $_ -ne "`r`n" }
	Create-Output-File
	$firstRun = $false
	Prompt-For-More-URLs
}

Main

Write-Host "Thanks for using autolink. Goodbye!"
