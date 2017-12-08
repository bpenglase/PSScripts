Try {
    $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $ComputerName=$tsenv.Value("OSDComputerName")
    if ($ComputerName -like "10-*") {
        $tsenv.Value("OSDComputerName")=$ComputerName.Replace("10-",'')
        $tsenv.Value("NHSDWinVer")="10"
    } elseif ($ComputerName -like "12-*") {
        $tsenv.Value("OSDComputerName")=$ComputerName.Replace("12-",'')
        $tsenv.Value("NHSDWinVer")="12"
    } elseif ($ComputerName -like "8-*") {
        $tsenv.Value("OSDComputerName")=$ComputerName.Replace("8-",'')
        $tsenv.Value("NHSDWinVer")="8"
    } elseif ($ComputerName -like "16-*") {
        $tsenv.Value("OSDComputerName")=$ComputerName.Replace("16-",'')
        $tsenv.Value("NHSDWinVer")="16"
    } elseif ($ComputerName -like "7-*") {
		# If for some reason 7 is still requested.. run with 8 instead...
		$tsenv.Value("OSDComputerName")=$ComputerName.Replace("7-",'')
		$tsenv.Value("NHSDWinVer")="8"
	}else {
        $tsenv.Value("NHSDWinVer") = "10"
    }
}
Catch {
    Write-Host "This only runs in a Task Sequence"
    echo "Failure" |  Out-File C:\ps.log 
}