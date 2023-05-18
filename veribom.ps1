
<#PSScriptInfo

.VERSION 0.0.13

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


$veribom_dir = Split-Path $MyInvocation.MyCommand.Path

$normal_number  = "(\d{5,8}\.?\w?(-?\d{0,2})?)"
$us_number      = "(US\d{4})"
$kit_number     = "(KIT ?#\d{1,4})"
$part_number    = "(?<!\()" + "(\(?([0-9]*\.?[0-9]+)[X'`"]\)?)?" + "(" + $normal_number + "|" + $us_number + "|" + $kit_number + ")" + "\n? ?(\(?([0-9]*\.?[0-9]+) ?[X'`"]\)?)?"
#                 not a ref#           leading quantity                             the main part number                               trailing quantity, maybe on the next line

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

function get_file($title) {
    # Get a file from the user
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
    $FileBrowser.Title = $title
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
    $csv = $csv | Select-Object -Property * -ExcludeProperty description,uom
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
    return $bom
}
function combine_boms($excel_bom, $pdf_bom) {
    $bom = @()
    # Add the excel parts to the bom
    foreach ($i in 0..($excel_bom.length-1)) {
        $bom += [pscustomobject]@{part_number=$excel_bom.part_number[$i];xls=[double]$excel_bom.quantity[$i];pdf=" "}
    }
    #Add the pdf parts to the bom
    foreach ($pdf_part_number in $pdf_bom.part_number) {
        $pdf_quantity = $pdf_bom.quantity[$pdf_bom.part_number.IndexOf($pdf_part_number)]
        if ($bom.part_number.contains($pdf_part_number)) {
            $loc = $bom.part_number.IndexOf($pdf_part_number)
            $bom[$loc].pdf = $pdf_quantity
        }
        else {
            $bom += [pscustomobject]@{part_number=$pdf_part_number;xls=" ";pdf=[double]$pdf_quantity}
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
        }
    return $bom 
}


# MAIN SCRIPT STARTS HERE
Clear-Host

# Write-Host "select an excel bom..." -NoNewline
# $xls_file = get_file -title "Select an Excel BoM"
# if ($xls_file -eq "") {return}
# Write-Host "done"
$xls_file = "C:\Users\apambrose\Documents\My_Drive\Projects\Powershell_Projects\veribom\more_test_data\B24058_D.xlsx"

# Write-Host "select a pdf drawing..." -NoNewline
# $pdf_file = get_file -title "Select a PDF BoM"
# if ($pdf_file -eq "") {return}
# Write-Host "done"
$pdf_file = "C:\Users\apambrose\Documents\My_Drive\Projects\Powershell_Projects\veribom\more_test_data\B24058_D.PDF"

$pdf_bom = pdf_to_bom $pdf_file
$xls_bom = excel_to_bom $xls_file

combine_boms -excel_bom $xls_bom -pdf_bom $pdf_bom 
