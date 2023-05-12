$veribom_dir = Split-Path $MyInvocation.MyCommand.Path
function pdf_to_text($pdf_path) {
    # Dont forget to unblock this guy during install
	Add-Type -Path "$veribom_dir\itextsharp.dll"
	$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $pdf_path
    [string]$text_out = @()
	for ($page = 1; $page -le $pdf.NumberOfPages; $page++){
		$text=[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
		$text_out += $text;
	}	
	$pdf.Close()
    return $text_out
}

function pdftext_to_bom($text) {
    $bom = @()
    $part_number_format = "((\d)X)? ?(\d{5,8}) ?((\d)X)?" # A 5-8 digit number that might have a leading/ trailing quantity
    foreach ($callout in (($text -split "`n") -match $part_number_format)) {
        $null = $callout -match $part_number_format
        $part_number = [string] $matches[3]
        $quantity_1  = [float]  $matches[2]
        $quantity_2  = [float]  $matches[5]
        $quantity = $quantity_1 + $quantity_2
        if ($quantity -eq 0){ $quantity = 1 }

        # If the part number already exists, update the quantity
        if ($bom -and $bom.part_number.contains($part_number)) {
            $index = $bom.IndexOf($part_number)
            $bom.quantity[$index] += $quantity
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
    $csv = $csv[3..($csv.length -1)]          # remove the header
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
        $bom += [pscustomobject]@{part_number=$excel_bom.part_number[$i];xls=$excel_bom.quantity[$i];pdf=" "}
    }
    #Add the pdf parts to the bom
    foreach ($pdf_part_number in $pdf_bom.part_number) {
        $pdf_quantity = $pdf_bom.quantity[$pdf_bom.part_number.IndexOf($pdf_part_number)]
        if ($bom.part_number.contains($pdf_part_number)) {
            $loc = $bom.part_number.IndexOf($pdf_part_number)
            $bom[$loc].pdf = $pdf_quantity
        }
        else {
            $bom += [pscustomobject]@{part_number=$pdf_part_number;xls=" ";pdf=$pdf_quantity}
        }
    }
    return $bom
}


# MAIN SCRIPT STARTS HERE

Write-Host "select an excel bom..." -NoNewline
$xls_file = get_file -title "Select an Excel BoM"
Write-Host "done"
Write-Host "select a pdf bom..." -NoNewline
$pdf_file = get_file -title "Select a PDF BoM"
Write-Host "done"

$pdf_bom = pdf_to_bom $pdf_file
$xls_bom = excel_to_bom $xls_file

combine_boms -excel_bom $xls_bom -pdf_bom $pdf_bom