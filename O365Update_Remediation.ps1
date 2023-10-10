# Written to run ClickToRun if Office is affected by CVE-2023-23397 - Outlook NTLM fla
$notUpdatedAfter=[datetime]"2023-03-15"


try {
    $C2R=get-item "$($env:ProgramFiles)\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
    if($c2r.CreationTime -lt $notUpdatedAfter) {
        # OK, so we need to get the current click2run process so we do not confuse the launched one with the preexisting one
        $initialProcesses=get-process -name "OfficeClickToRun"
        write-output "Spawning update as OfficeC2RClient ($($c2r.versioninfo.FileVersion)) is old with a date of $($c2r.creationtime)"
        $newProcess=start-process -FilePath "$($env:ProgramFiles)\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" -argumentlist "/update user forceappshutdown=false displaylevel=true" -passthru -wait
        # sleep a little to allow it to start
        start-sleep -seconds 60
        # Get the new process (if it exists)
        $spawnedProcess=(get-process -name "OfficeClickToRun"|where {-not($_.id -in $initialProcesses.id)})
        if ($spawnedProcess) {
            wait-process $spawnedProcess.id
        }
        # Keeps being reported as recurring in proactive remediations - maybe we're exiting too fast?
        start-sleep -seconds 60 
        $C2RAfter=get-item "$($env:ProgramFiles)\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
        write-output "Post Update: OfficeC2RClient has version of ($($c2rafter.versioninfo.FileVersion)) with a date of $($c2rafter.creationtime)"


        exit 0
    }
    else {
        write-output "All Good: OfficeC2RClient ($($c2r.versioninfo.FileVersion)) has a date of $($c2r.creationtime)"
        exit 0
    }
}
catch {
    write-output "Office Click2Run not found!"
    exit 0
}