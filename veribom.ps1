
<#PSScriptInfo

.VERSION 0.5.47

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

$ignored_number = "(BEGA)|(NOTE)|(ASM)|((\b4)(?!(2885)|(2886)))"

$normal_number  = "(\d{5,8}\.?\w?(-?\d{0,2})?)"
$us_number      = "(US\d{4})"
$kit_number     = "(KIT ?#\d{1,4})"
$any_number     = "(" + $normal_number + "|" + $us_number + "|" + $kit_number + ")"
$leading_qty    = "(?<!\()" + "(\(?([0-9]*\.?[0-9]+)[X'`"]\)? ?)"
$trailing_qty   = "\n?( ?\(?([0-9]*\.?[0-9]+) ?[X'`"]\)?)"
$part_number    = "(" + $leading_qty + $any_number + ")|(" + $any_number + $trailing_qty + ")|(" + $any_number + ")"
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
            Write-Host "would you like to update? (y/n)"
            $choice = [Console]::ReadKey("No Echo").KeyChar
            if ($choice -eq "y") {
                Get-ChildItem $veribom_dir | Remove-Item -Recurse -ErrorAction SilentlyContinue
                Get-ChildItem "$home\downloads\veribom_temp" | Move-Item -Destination $veribom_dir
                Write-Host "update complete!"
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
    $max_line_length = 38
    $text = $text -split "`n" 
    $text = $text.where({
        ($_.length -lt $max_line_length) -and `
        (-not($_ -match "REPLACED")) -and `
        (-not($_ -match "REMOVED")) -and `
        (-not($_ -match "ADDED")) -and `
        (-not($_ -match "CHANGED"))
    })
    $text = $text -join "`n"

    $bom = @()
    $callouts = $text | Select-String -Pattern $part_number -AllMatches
    foreach ($callout in $callouts.Matches) {
            if (($part_number = [string] $callout.Groups[4].Value))  {}
        elseif (($part_number = [string] $callout.Groups[10].Value)) {}
        elseif (($part_number = [string] $callout.Groups[18].Value)) {}
            if (($quantity    = [double] $callout.Groups[3].Value))  {}
        elseif (($quantity    = [double] $callout.Groups[16].Value)) {}

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
    $csv = $csv.where({$_.part_number -ne ""})      # remove empty elements (including leading elements)
    if ($csv.part_number.IndexOf("Part Number") -ge 0) {
        $bom = $csv[($csv.part_number.IndexOf("Part Number")+1)..($csv.part_number.length -1)] 
    } elseif ($csv.part_number.IndexOf("Component") -ge 0) {
        $bom = $csv[($csv.part_number.IndexOf("Component")+1)..($csv.part_number.length -1)]
    } else {
        throw "bom does not start with a 'part number' or 'component' column"
    }
       # remove the header
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



function combine_boms {

    param(
        $excel_bom,
        $pdf_bom
    )

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

    return $bom
}

function red($string) {
    $e = [char]27
    $color = 31
    "$e[${color}m$($string)${e}[0m"
}
function yellow($string) {
    $e = [char]27
    $color = 33
    "$e[${color}m$($string)${e}[0m"
}
function format_row ($row, $val) {
    if (($row.pdf -eq $row.xls)) {
        $val
    } elseif ($row.part_number -match $ignored_number) {
        yellow $val
    } else {
        red $val
    }
}

function display_bom {
    param(
        $bom,
        [boolean] $display_errors   = $false,
        [boolean] $display_warnings = $false,
        [boolean] $display_ok       = $false
    )

    if (-not $display_errors){
        $bom = $bom.where({($_.part_number -match $ignored_number) -or ($_.pdf -eq $_.xls)})
    }
    if (-not $display_warnings) {
        $bom = $bom.where({-not ($_.part_number -match $ignored_number)})
    }
    if (-not $display_ok) {
        $bom = $bom.where({($_.pdf -ne $_.xls)})
    }
    
    $bom = $bom | Sort-Object part_number
    $bom = $bom | Format-Table `
        @{Name='Part Number'; Align="left";  Expression={format_row $_ $_.part_number}},`
        @{Name='XLS';         Align="right"; Expression={format_row $_ $_.xls}}, `
        @{Name='PDF';         Align="left";  Expression={format_row $_ $_.pdf}}, `
        @{Name='Description'; Align="left";  Expression={format_row $_ $_.description}}

    return $bom
}

$global:starting_directory = "$home"
function new_comparison () {

    param(
        [switch] $full = $false
    )

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

    $combined_bom = combine_boms -excel_bom $xls_bom -pdf_bom $pdf_bom

    $display_errors   = $true
    $display_warnings = $false
    $display_ok       = $false

    if ($full) {
        $display_errors   = $true
        $display_warnings = $true
        $display_ok       = $true
    }

    

    :process_bom while($true) {
        display_bom -bom $combined_bom -display_errors $display_errors -display_warnings $display_warnings -display_ok $display_ok
        
        Write-Host "use " -NoNewline
        Write-Host "e" -ForegroundColor "yellow" -NoNewline
        Write-Host ", " -NoNewline
        Write-Host "w" -ForegroundColor "yellow" -NoNewline
        Write-Host ", and " -NoNewline
        Write-Host "r" -ForegroundColor "yellow" -NoNewline
        Write-Host " to toggle erros, warnings, and ok parts"
        Write-Host "press " -NoNewline
        Write-Host "q" -ForegroundColor "Yellow" -NoNewline
        Write-Host " to quit"

        :valid_command while($true) {
            $command = [Console]::ReadKey("No Echo").KeyChar
            switch ($command) {
                "e"     {$display_errors   = -not $display_errors}
                "w"     {$display_warnings = -not $display_warnings}
                "r"     {$display_ok       = -not $display_ok}
                "q"     {break process_bom}
                default {continue valid_command}
            }
            break
        }
        Clear-Host
    }
}

############## The main script starts here
[Console]::CursorVisible = $false
Clear-Host
if ($no_update -eq $false){
check_for_updates
}

function convert_single_pdf {
    Write-Host "select a drawing: " -NoNewline
    $pdf_file = get_file -title "Select A Drawing PDF" -starting_dir $starting_directory -filter "Drawing (*.pdf)|*.pdf"
    if ($pdf_file -eq "") {""; return}
    Split-Path $pdf_file -Leaf
    $global:starting_directory = Split-Path $pdf_file -Parent
    ""
    pdf_to_text $pdf_file
}

$show_hidden_commands = $false
# Main Loop
:main while ($true) {
    Clear-Host
    Write-Host ("---------    veribom " + $veribom_ver.Major + "." + $veribom_ver.Minor + "     ---------")
    Write-Host "n" -ForegroundColor "Yellow" -NoNewline; ")  new/next veribom"
    Write-Host "h" -ForegroundColor "Yellow" -NoNewline; ")  help, open the veribom project page"
    Write-Host "u" -ForegroundColor "Yellow" -NoNewline; ")  update/ check for updates"
    Write-Host "t" -ForegroundColor "Yellow" -NoNewline; ")  toggle hidden commands"
    if ($show_hidden_commands){
        Write-Host "r" -ForegroundColor "Yellow" -NoNewline; ")  raw pdf, see what veribom it looking at"
        Write-Host "e" -ForegroundColor "Yellow" -NoNewline; ")  regex used for part number matching"
        Write-Host "v" -ForegroundColor "Yellow" -NoNewline; ")  version of veribom"
        Write-Host "x" -ForegroundColor "Yellow" -NoNewline; ")  exit the veribom program"
    }

    # Keep looping until we get one of the available commands
    :valid_command while($true) {
        $command = [Console]::ReadKey("No Echo").KeyChar
        switch ($command) {
            "n"     {try{new_comparison; continue main}catch{Write-Host $_ -ForegroundColor "red"}}
            "h"     {Start-Process "https://github.com/AustinPAmbrose/veribom"}
            "u"     {"";check_for_updates}
            "t"     {$show_hidden_commands = -not $show_hidden_commands; continue main}
            "r"     {convert_single_pdf}
            "e"     {"";$part_number}
            "v"     {"";"$veribom_ver"}
            "x"     {return}
            default {continue valid_command}
        }
        break
    }
        ""
        Write-Host "press" -NoNewline
        Write-Host " q " -NoNewline -ForegroundColor "Yellow"
        Write-Host "to quit"
        while([Console]::ReadKey("No Echo").KeyChar -ne "q"){}
}
