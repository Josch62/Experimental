<#
    Author: Johan Schrewelius
    Name:   DismCapture.ps1
    Usage:  Script that runs dism to Capture a windows image. Run as >>Run Command<< step from Package (will fail otherwise). Script expects no parameters.
    Dependencies: All of the following TS variables must exist and contain valid data: %CaptureDir%, %ImageFile%, %Compression%
    Optional: %ConfigFile%"
    Version: 0.8
    Date: 2020-09-09
#>

begin {
    # Create Com objects
    $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $tsUI = New-Object -COMObject Microsoft.SMS.TsProgressUI

    # Get percentage as double
    function Get-Percentage ($Text) {
        $str = "$Text" -replace "[^0-9.]" , ''
        [double]$val = 0.0
        [double]::TryParse($str, [ref]$val) | Out-Null
        $val
    }

    # Get TS Variables
    $CaptureDir = $tsenv.Value("CaptureDir").TrimEnd('\')
    $ImageFile = $tsenv.Value("ImageFile")
    $SMSTSMachineName = $tsenv.Value("_SMSTSMachineName")
    $Compression = $tsenv.Value("Compression")
    $ConfigFile = $tsenv.Value("ConfigFile")
    $LogPath = $tsenv.Value("_SMSTSLogPath")
    $OrgName = $tsenv.Value("_SMSTSOrgName")
    $TSName = $tsenv.Value("_SMSTSPackageName")
    $CurrStepName = $tsenv.Value("_SMSTSCurrentActionName")
    [int]$CurrentStep = [int]$tsenv.Value("_SMSTSNextInstructionPointer")
	[int]$TotalSteps = [int]$tsenv.Value("_SMSTSInstructionTableSize")

	If ([string]::IsNullOrEmpty($tsenv.Value("TSProgressInfoLevel"))) {
		$TSProgressInfoLevel = $null
	}
	else {
		[int]$TSProgressInfoLevel = [int]$tsenv.Value("TSProgressInfoLevel")
	}

    # Declare and define variables

    $Self = "DismCapture.ps1"
    $LogFile = "$LogPath\Onevinn." + $Self.Replace(".ps1", ".log")

    # Log in CMTrace format
    function WriteLog {
        param(
        [Parameter(Mandatory)]
        [string]$LogText,
        [Parameter(Mandatory=$true)]
        $Component,
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Info','Warning','Error','Verbose')]
        [string]$Type,
        [Parameter(Mandatory)]
        [string]$LogFileName,
        [Parameter(Mandatory)]
        [string]$FileName
        )

        switch ($Type)
        {
            "Info"      { $typeint = 1 }
            "Warning"   { $typeint = 2 }
            "Error"     { $typeint = 3 }
            "Verbose"   { $typeint = 4 }
        }

        $time = Get-Date -f "HH:mm:ss.ffffff"
        $date = Get-Date -f "MM-dd-yyyy"
        $ParsedLog = "<![LOG[$($LogText)]LOG]!><time=`"$($time)`" date=`"$($date)`" component=`"$($Component)`" context=`"`" type=`"$($typeint)`" thread=`"$($pid)`" file=`"$($FileName)`">"
        $ParsedLog | Out-File -FilePath "$LogFileName" -Append -Encoding utf8
    }

    # Assemble dism parameters
    $Parameters = "/Capture-Image /CaptureDir:`"$($CaptureDir)`" /ImageFile:`"$($ImageFile)`" /Name:`"$($SMSTSMachineName)`" /Compress:`"$($Compression)`"" 
    
    if ($ConfigFile) {
        $Parameters += " /ConfigFile:`"$($ConfigFile)`""
    }

    # Helper Function to feed ShowActionProgress, less arguments
    function ShowProgress ([int]$Percentage){

        if ($null -eq $TSProgressInfoLevel) {
            $tsUI.ShowActionProgress("$OrgName", "$TSName", $null, "$CurrStepName", $CurrentStep, $TotalSteps, "Capture $($Percentage)% done", $Percentage, 100)
        }
        else {
            $tsUI.ShowActionDetailedProgress("$OrgName", "$TSName", $null, "$CurrStepName", $CurrentStep, $TotalSteps, "Dism Image Capture", $Percentage, 100, $TSProgressInfoLevel)
        }
    }

    # Setup stdin\stdout redirection
    $StartInfo = New-Object System.Diagnostics.ProcessStartInfo -Property @{
                    FileName = "Dism.exe"
                    Arguments = "$Parameters"
                    UseShellExecute = $false
                    RedirectStandardOutput = $true
                    RedirectStandardError = $true
                    CreateNoWindow = $true
                    WorkingDirectory = "$env:TEMP"
                }

    # Create new process
    $Process = New-Object System.Diagnostics.Process

    # Assign previously created StartInfo properties
    $Process.StartInfo = $StartInfo
    $Process.EnableRaisingEvents = $true

    # Register Object Events for stdin\stdout reading
    $OutEvent = Register-ObjectEvent -InputObject $Process -EventName OutputDataReceived -Action {

        if (![string]::IsNullOrEmpty($Event.SourceEventArgs.Data)) {

            [string]$text = ($Event.SourceEventArgs.Data).Trim()

            WriteLog -LogFileName "$LogFile" -Component "RunPowerShellScript" -FileName "$Self" -LogText "$text" -Type Info

            if ($text -match "\[([^\[]*)\]") {
                ShowProgress -Percentage (Get-Percentage -Text $text)
            }
        }
    }

    $ErrEvent = Register-ObjectEvent -InputObject $Process -EventName ErrorDataReceived -Action {
        
        if (![string]::IsNullOrEmpty($Event.SourceEventArgs.Data)) {

            [string]$text = ($Event.SourceEventArgs.Data).Trim()
            WriteLog -LogFileName "$LogFile" -Component "RunPowerShellScript" -FileName "$Self" -LogText "$text" -Type Error
        }

    }

    $ExitEvent = Register-ObjectEvent -InputObject $Process -EventName Exited -Action {
        New-Event -SourceIdentifier "TimeToExit"
    }
}

process {
    # Start process
    $Process.Start() | Out-Null

    # Begin reading stdin\stdout
    $Process.BeginOutputReadLine()
    $Process.BeginErrorReadLine()

    # Wait for process exit
    Wait-Event -SourceIdentifier "TimeToExit" -ErrorAction SilentlyContinue | Out-Null
    Remove-Event -SourceIdentifier "TimeToExit" -ErrorAction SilentlyContinue | Out-Null
}

end {

    # Unregister events
    $OutEvent.Name, $ErrEvent.Name, $ExitEvent.Name | ForEach-Object { Unregister-Event -SourceIdentifier $_ }

    # Exit with dism return code
    $Process.ExitCode
}
