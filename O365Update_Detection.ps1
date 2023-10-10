# Written to detect if Office is affected by CVE-2023-23397 - Outlook NTLM fla
# Validate whether Click2Run.exe has a modifed date >= 15/03/2023
#
# We cannot easily check office version numbers as the patched version differs across multiple
# channel versions, and is not numerically subsequent to unpatched versions (aside from within that channel)

$notUpdatedAfter=[datetime]"2023-03-15"

try {
    $c2r=get-item "$($Env:ProgramFiles)\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"
    if($c2r.CreationTime -lt $notUpdatedAfter) {
        write-output "OUTOFDATE: OfficeC2RClient ($($c2r.versioninfo.FileVersion)) is old with a date of $($c2r.creationtime)"
        exit 1
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