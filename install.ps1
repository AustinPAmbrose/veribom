# Make sure their powershell version is at least 5.1
if ($PSVersionTable.PSVersion -lt [version]::new("5.1")) {
    Write-Warning "this script was written for powershell version 5.1,"
    Write-Warning ("but your current powershell version is " + $PSVersionTable.PSVersion)
    Write-Warning "see this link for instructions on updating your powershell:"
    "https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.3"
    return
}

# Make sure they have the correct permissions to execute scripts
$policy_problem = switch (Get-ExecutionPolicy) {
    Restricted   {$true}
    AllSigned    {$true}
    Bypass       {$false}
    RemoteSigned {$false}
    Undefined    {$false}
    Unrestricted {$false}
    default      {$false}
}
if ($policy_problem) {
    Write-Warning "you may not be able to run scripts with your current script execution policy"
    Write-Warning "please review the following microsoft documentation on execution policies:"
    "https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.3"
    Write-Host "Would you like to update your execution policy to RemoteSigned? (y/n)"
    $do_update_policy = [Console]::ReadKey("No Echo").KeyChar
    if ($do_update_policy -eq "y") {
        Set-ExecutionPolicy RemoteSigned -Scope Process
        Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
        if ((Get-ExecutionPolicy) -eq [Microsoft.PowerShell.ExecutionPolicy]::RemoteSigned) {
            "execution policy successfully updated to RemoteSigned"
        } else {
            "execution policy failed to update"
            "the installer will now exit"
            return
        }
    } else {
        "your execution policy will not be updated, but this installation will now exit"
        return
    }
}

# Now for the actual downloading
Clear-Host
[console]::CursorVisible = $false
$ProgressPreference = "SilentlyContinue"
Write-Host "veribom will be installed to --> $home\.veribom"
Write-Host "downloading..." -NoNewline
Invoke-WebRequest "https://github.com/AustinPAmbrose/veribom/raw/main/release.zip" -OutFile "$home\downloads\release.zip"
Write-Host "done!"
Write-Host "installing..." -NoNewline
Expand-Archive "$home\downloads\release.zip" -DestinationPath "$home\.veribom" -Force
Remove-Item "$home\downloads\release.zip"
Write-Host "done!"

Write-Host "creating desktop shortcut..." -NoNewline
$WshShell = New-Object -comObject WScript.Shell
$shortcut = $WshShell.CreateShortcut("$Home\Desktop\veribom.lnk")
$shortcut.TargetPath = "powershell.exe"
$shortcut.Arguments  = "$home\.veribom\veribom.ps1"
$shortcut.Hotkey     = "CTRL+ALT+A"
$shortcut.Save()
Write-Host "done!"

""
"veribom installed successfully!"
"you may now close this window"
""
"use `'CTRL + ALT + A`' to launch veribom..."
while ($true) {}