
param(
    [Parameter(Mandatory=$true)][string]$command,
    [string]$port = "",
    [switch]$log = $false
)

$scriptName = $MyInvocation.MyCommand.Name

Function Main {
    if ([string]::IsNullOrWhiteSpace($port)) {
        $ports = [System.IO.Ports.SerialPort]::getportnames()
        Log "Found ports $($ports)"

        switch ($ports.Length)
        {
            0 {
                Log "No COM port found. Exiting.";
                Exit;
            }

            1 {
                $port = $ports[0];
                Log "Using port $($port)"
                Break;
            }

            default {
                Log "More than one COM port found. Specify which is connected to the projector with -port switch. Exiting."
                Exit;
            }
        }
    }

    # Check if this script is already running
    WaitForOthers


    $comPort = New-Object System.IO.Ports.SerialPort $port,9600,None,8,one
    $comPort.Open();

    switch ($command) {
        "lens1" {
            #$comPort.WriteLine("$([char] 2)VXX:LMLI0=+00000$([char] 3)");
            SetLensMemory 0
            Break;
        }
        "lens2" {
            SetLensMemory 1
            Break;
        }
        "lens3" {
            SetLensMemory 2
            Break;
        }
        "lens4" {
            SetLensMemory 3
            Break;
        }
        "lens5" {
            SetLensMemory 4
            Break;
        }
        "lens6" {
            SetLensMemory 5
            Break;
        }
    }

    $comPort.Close();
    $comPort.Dispose();
}


Function SetLensMemory {
    param(
        [Parameter(Mandatory=$true)][int32]$number
    )

    Log "Sending lens$($number + 1)"
    $comPort.WriteLine("$([char] 2)VXX:LMLI0=+0000$($number)$([char] 3)");

    # Pause before exiting to give projector time to focus - this mechanism is used to help us queue lens commands
    Start-Sleep -Seconds 8
}


Function WaitForOthers {
    [boolean]$shouldWait = $true
    Log "Our process id is $($PID)"

    while ($shouldWait) {

        # Find all scripts running with our name and sort by execution time. 
        $process = Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe' AND CommandLine LIKE '%$scriptName%'" | Sort-Object -Property UserModeTime -Descending | Select-Object -First 1
        Log "Oldest process is $($process.ProcessId)"
        
        # If we are not the oldest, keep waiting
        if ($process.ProcessId -ne $PID)
        {
            Log "Which isn't us. Sleeping."
            Start-Sleep -Milliseconds 500
        }
        else {
            Log "Which is us. Continuing."
            $shouldWait = $false
        }
    }
}


Function Log {
    param(
        [Parameter(Mandatory=$true)][String]$message
    )

    Write-Output $message

    if ($log) {
        Add-Content "$([environment]::GetFolderPath('MyDocuments'))\$($scriptName).log" $message
    }
}

Main