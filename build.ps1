# Make sure we're in the right dirrectory
if (-not (Test-Path ./veribom.ps1)) {
    throw "you're not in the veribom directory :/"
    return
}

# Make sure the script doesn't contain any errors
if ((Invoke-ScriptAnalyzer ./veribom.ps1).Severity -join "" -match "Error") {
    throw "script contains errors..."
    return
}

# Check that everything is committed and pushed
$null = git fetch --quiet
$status = git status
if ((-not $status.contains("On branch main")) -or `
   ($status.contains("Changes not staged for commit")) -or `
   ($status.contains("Your branch is behind 'origin/main'"))) {
    throw "git error... pull, commit, push, and try again"
    return
}

# If we made it here, the script is error-free and up-to-date

# Ask the dev what version they want to build
$v          = (Test-ScriptFileInfo ./veribom.ps1).Version
$next_major = [version]::New($v.Major+1,    0      ,$v.Build+1)
$next_minor = [version]::New($v.Major  ,$v.Minor+1 ,$v.Build+1)
$next_build = [version]::New($v.Major  ,$v.Minor   ,$v.Build+1)
"What version would you like to build?"
"    1) Major Version - " + $next_major
"    2) Minor Version - " + $next_minor
"    3) Build Version - " + $next_build
switch (Read-Host) {
    "1" {$next_version = $next_major}
    "2" {$next_version = $next_minor}
    "3" {$next_version = $next_build}
    default {
        throw "choice must be 1, 2, or 3"
        return
    }
}
Update-ScriptFileInfo ./veribom.ps1 -Version $next_version

# The script is updated now. Move everything we need into a zip file
Copy-Item -Path ./veribom.ps1, ./itextsharp.dll `
          -Destination ./release
Compress-Archive -Path ./release -DestinationPath ./release.zip -Force
Remove-Item ./release -Recurse -Force

# Push that mofo to origin
$null = git add "./release.zip"
$null = git commit -am $next_version.ToString() --quiet
$null = git push --quiet
""
"done!"