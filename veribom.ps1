
<#PSScriptInfo

.VERSION 0.2.41

.GUID 1eb7878d-24c4-4677-87b7-478a7502bd37

.AUTHOR aambrose

.COMPANYNAME 

.COPYRIGHT 

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES


#> 

<# 

.DESCRIPTION 
 automate the bom checking process! 

#> 

param([Switch]$no_update = $false)

$veribom_loc = $MyInvocation.MyCommand.Path
$veribom_dir = Split-Path $veribom_loc -Parent
$veribom_ver = (Test-ScriptFileInfo $veribom_loc).Version

$normal_number  = "(\d{5,8}\.?\w?(-?\d{0,2})?)"
$us_number      = "(US\d{4})"
$kit_number     = "(KIT ?#\d{1,4})"
$part_number    = "(?<!\()" + "(\(?([0-9]*\.?[0-9]+)[X'`"]\)?)? ?" + "(" + $normal_number + "|" + $us_number + "|" + $kit_number + ")" + "\n? ?(\(?([0-9]*\.?[0-9]+) ?[X'`"]\)?)?"
#                 not a ref#           leading quantity                             the main part number                               trailing quantity, maybe on the next line

function check_for_updates {
    try {
        [console]::CursorVisible = $false
        $download = Start-Job -ScriptBlock {
            try {
                $ProgressPreference = "SilentlyContinue"
                Invoke-WebRequest "https://github.com/AustinPAmbrose/veribom/raw/main/release.zip" -OutFile "$home\downloads\veribom_temp.zip"
                Remove-Item "$home\downloads\veribom_temp" -Recurse -Force -ErrorAction SilentlyContinue
                Expand-Archive "$home\downloads\veribom_temp.zip" -DestinationPath "$home\downloads\veribom_temp" -Force
                Remove-Item "$home\downloads\veribom_temp.zip" -ErrorAction SilentlyContinue
                $next_version = (Test-ScriptFileInfo "$home\downloads\veribom_temp\veribom.ps1").Version
                return $next_version
            } catch {
                throw $_
            }
        }
        while ($download.State -eq "NotStarted") {}
        while ($download.State -eq "Running") {
            Write-Host "`rchecking for updates   " -NoNewline; Start-Sleep -Milliseconds 200
            Write-Host "`rchecking for updates.  " -NoNewline; Start-Sleep -Milliseconds 200
            Write-Host "`rchecking for updates.. " -NoNewline; Start-Sleep -Milliseconds 200
            Write-Host "`rchecking for updates..." -NoNewline; Start-Sleep -Milliseconds 200
        }
        $null = Wait-Job $download
        if ($download.State -eq "Failed") {throw $download.JobStateInfo.Reason.Message}
        $new_version = Receive-Job $download
        Remove-Job $download

        if ($new_version -gt $veribom_ver) {
            [console]::CursorVisible = $true
            Write-Host "new update available!"
            Write-Host "version " -NoNewline
            Write-Host $veribom_ver.ToString() -ForegroundColor "Yellow" -NoNewline
            Write-Host " -> " -NoNewline
            Write-Host $new_version.ToString() -ForegroundColor "Yellow"
            Write-Host ""
            Write-Host "ChangeLog:"
            $ProgressPreference = "SilentlyContinue"
            $changelog = Invoke-WebRequest "https://github.com/AustinPAmbrose/veribom/raw/main/CHANGELOG.md"
            foreach ($note in ($changelog.Content -split "`n")) {
                if ($note -match $veribom_ver.ToString()){break}
                if ($note -match "\d+.\d+.\d+") {
                    Write-Host $note -ForegroundColor "Yellow"
                } else {
                    Write-Host "   -$note"
                }
            }
            Write-Host ""
            Write-Host"would you like to update? (y/n)"
            $choice = [Console]::ReadKey("No Echo").KeyChar
            if ($choice -eq "y") {
                Get-ChildItem $veribom_dir | Remove-Item -Recurse -ErrorAction SilentlyContinue
                Get-ChildItem "$home\downloads\veribom_temp" | Move-Item -Destination $veribom_dir
                Write-Host "update vomplete!"
                powershell $veribom_loc -no_update
                while($true) {}
            }
        }
    } catch {
        Write-Host "failed to check for updates"
        Write-Host "$_"
        Start-Sleep -Seconds 1
    } finally {
        Remove-Item "$home\downloads\veribom_temp.zip" -ErrorAction SilentlyContinue
        Remove-Item "$home\downloads\veribom_temp" -ErrorAction SilentlyContinue -Recurse
        Clear-Host
    }
}

function pdf_to_text($pdf_path) {
    # Dont forget to unblock this guy during install
	Add-Type -Path "$veribom_dir\itextsharp.dll"
	$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $pdf_path
    $page = 1
	$text=[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
	$pdf.Close()
    return $text
}

function pdftext_to_bom($text) {

    # Get rid of all lines that are longer than max_line_length
    $max_line_length = 25
    $text = $text -split "`n" 
    $text = $text.where({$_.length -lt $max_line_length})
    $text = $text -join "`n"

    $bom = @()
    $callouts = $text | Select-String -Pattern $part_number -AllMatches
    foreach ($callout in $callouts.Matches) {
        $part_number = [string] $callout.Groups[3].Value
        $quantity_1  = [double] $callout.Groups[2].Value
        $quantity_2  = [double] $callout.Groups[9].Value
        $quantity = $quantity_1 + $quantity_2
        if ($quantity -eq 0){ $quantity = 1 }

        # If the part number already exists, update the quantity
        if ($bom -and $bom.part_number.contains($part_number)) {
            $index = $bom.part_number.IndexOf($part_number)
            $bom[$index].quantity += $quantity
        } else {
            $bom += [pscustomobject]@{part_number=$part_number;quantity=$quantity}
        }
    }
    return $bom
}

function get_file($title, $starting_dir, $filter) {
    # Get a file from the user
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
    $FileBrowser.Title            = $title
    $FileBrowser.InitialDirectory = $starting_dir
    $FileBrowser.Filter           = $filter
    $null = $FileBrowser.ShowDialog()
    return $FileBrowser.FileName
}

function excel_to_csv($xls_path, $csv_path) {
    # Convert the excel to a csv
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($xls_path)
    $wb.SaveAs($csv_path, 6) 
    $E.Quit()
}

function csv_to_bom ($csv_path) {
    # Returns a table of parts 
    $csv = Import-Csv $csv_path -Header "part_number", "description", "uom", "quantity"
    $csv = $csv[3..($csv.length -1)]           # remove the header
    $bom = $csv.where({$_.part_number -ne ""}) # remove empty elements
    return $bom 
}

function pdf_to_bom($pdf_path) {
    $pdf_text = pdf_to_text -pdf_path $pdf_path
    $bom = pdftext_to_bom -text $pdf_text
    return $bom
}

function excel_to_bom($excel_path) {
    $csv_path = "$veribom_dir\temp.csv"
    excel_to_csv -xls_path $xls_file -csv_path $csv_path
    $bom = csv_to_bom -csv_path $csv_path
    Remove-Item "$veribom_dir\temp.csv" -ErrorAction SilentlyContinue
    return $bom
}
function combine_boms($excel_bom, $pdf_bom) {
    $bom = @()
    # Add the excel parts to the bom
    foreach ($i in 0..($excel_bom.length-1)) {
        $bom += [pscustomobject]@{part_number=$excel_bom.part_number[$i];description=$excel_bom.description[$i];xls=[double]$excel_bom.quantity[$i];pdf=$null}
    }
    #Add the pdf parts to the bom
    foreach ($pdf_part_number in $pdf_bom.part_number) {
        $pdf_quantity = $pdf_bom.quantity[$pdf_bom.part_number.IndexOf($pdf_part_number)]
        if ($bom.part_number.contains($pdf_part_number)) {
            $loc = $bom.part_number.IndexOf($pdf_part_number)
            $bom[$loc].pdf = $pdf_quantity
        }
        else {
            $bom += [pscustomobject]@{part_number=$pdf_part_number;description="";xls=" ";pdf=[double]$pdf_quantity}
        }
    }
    $bom = $bom | Sort-Object part_number
    $bom = $bom | Format-Table `
        @{
            Name='Part Number'
            Align="left"
            Expression={
                if ($_.pdf -eq $_.xls) {
                    $color = "0"
                } else {
                    $color = "31"
                }
                $e = [char]27                    
                "$e[${color}m$($_.part_number)${e}[0m"
            }
        }, `
        @{
            Name='XLS'
            Align="right"
            Expression={
                if ($_.pdf -eq $_.xls) {
                    $color = "0"
                } else {
                    $color = "31"
                }
                $e = [char]27                    
                "$e[${color}m$($_.xls)${e}[0m"
            }
        }, `
        @{
            Name='PDF'
            Align="left"
            Expression={
                if ($_.pdf -eq $_.xls) {
                    $color = "0"
                } else {
                    $color = "31"
                }
                $e = [char]27                    
                "$e[${color}m$($_.pdf)${e}[0m"
            }
        }, `
        @{
            Name='Description'
            Align="left"
            Expression={
                if ($_.pdf -eq $_.xls) {
                    $color = "0"
                } else {
                    $color = "31"
                }
                $e = [char]27                    
                "$e[${color}m$($_.description)${e}[0m"
            }
        }
    return $bom 
}

$global:starting_directory = "$home"
function new_comparison () {
    Clear-Host
    Write-Host "select a bom:     " -NoNewline 
    $xls_file = get_file -title "Select an Excel BoM" -starting_dir $starting_directory -filter "BoM (*.xlsx) |*.xlsx"
    if ($xls_file -eq "") {""; return}
    Split-Path $xls_file -Leaf
    $global:starting_directory = Split-Path $xls_file -Parent
    #$xls_file = "C:\Users\apambrose\Documents\My_Drive\Projects\Powershell_Projects\veribom\more_test_data\B24058_D.xlsx"

    Write-Host "select a drawing: " -NoNewline
    $pdf_file = get_file -title "Select A Drawing PDF" -starting_dir $starting_directory -filter "Drawing (*.pdf)|*.pdf"
    if ($pdf_file -eq "") {""; return}
    Split-Path $pdf_file -Leaf
    $global:starting_directory = Split-Path $pdf_file -Parent
    #$pdf_file = "C:\Users\apambrose\Documents\My_Drive\Projects\Powershell_Projects\veribom\more_test_data\B24058_D.PDF"

    $pdf_bom = pdf_to_bom $pdf_file
    $xls_bom = excel_to_bom $xls_file

    combine_boms -excel_bom $xls_bom -pdf_bom $pdf_bom 
}

############## The main script starts here
[Console]::CursorVisible = $false
Clear-Host
if ($no_update -eq $false){
check_for_updates
}

# Main Loop
:main while ($true) {
    Write-Host ("---------    veribom " + $veribom_ver.Major + "." + $veribom_ver.Minor + "     ---------")
    Write-Host "n" -ForegroundColor "Yellow" -NoNewline; ")  new/next veribom"
    Write-Host "h" -ForegroundColor "Yellow" -NoNewline; ")  help, open the veribom project page"
    Write-Host "v" -ForegroundColor "Yellow" -NoNewline; ")  version of veribom"
    Write-Host "u" -ForegroundColor "Yellow" -NoNewline; ")  update/ check for updates"
    Write-Host "x" -ForegroundColor "Yellow" -NoNewline; ")  exit the veribom program"

    # Keep looping until we get one of the available commands
    :valid_command while($true) {
        $command = [Console]::ReadKey("No Echo").KeyChar
        switch ($command) {
            "n"     {new_comparison}
            "h"     {Start-Process "https://github.com/AustinPAmbrose/veribom"}
            "v"     {"";"$veribom_ver"}
            "u"     {"";check_for_updates}
            "x"     {return}
            default {continue valid_command}
        }
        break
    }
        ""
        "press q to return to the main menu..."
        while([Console]::ReadKey("No Echo").KeyChar -ne "q"){}
        Clear-Host
}
