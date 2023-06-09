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
if ((-not ($status -match "On branch main")) -or `
   ($status -match "Changes not staged for commit") -or `
   ($status -match "Your branch is behind 'origin/main'")) {
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
switch ([Console]::ReadKey("No Echo").KeyChar) {
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
$items_to_release = @("./veribom.ps1", "./itextsharp.dll")
New-Item -Path./temp_dir -ItemType Directory -ErrorAction SilentlyContinue
Copy-Item -Path $items_to_release -Destination ./temp_dir
Get-ChildItem ./temp_dir -Recurse | Compress-Archive -DestinationPath ./release.zip -Force
Remove-Item ./temp_dir -Recurse -Force

# Update the changelog
$next_version.ToString() | Out-File ./CHANGELOG.md -Encoding ASCII
git log --pretty="%s" | Out-File -Append ./CHANGELOG.md -Encoding ASCII

# Push that mofo to origin
git commit -am $next_version.ToString() --quiet
git push --quiet
""
"done!"