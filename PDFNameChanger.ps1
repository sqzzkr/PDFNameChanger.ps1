Import-module PSexcel
$user = $([Environment]::UserName)
Add-Type -Path ".\itextsharp.dll"
$path = #path to folder
$pdfs = get-childitem -path $path -filter *pdf
$Excel = New-Excel -Path #path to xlsx file (in my case: firstname, lastname and location)
$WorkSheet = $Excel | Get-WorkSheet
$rows = $Worksheet.dimension.rows

for ($i = 2; $i -le $rows; $i++){
        $firstname = $WorkSheet.Cells[$i, 1].Text
        $lastname = $WorkSheet.Cells[$i, 2].Text
        $location = $WorkSheet.Cells[$i, 3].Text
        $toFilename = $lastname
        if($lastname.length -gt 30) {
        $lastname = $lastname.substring(0,30)
        }
        
        if ($location -eq "SomeLocation")
            {
	            $city = "SomeCity"
            }
		    elseif ($location -eq "AnotherLocation")
            {
	            $city = "AnotherCity"
            }
		    else{
				$city = "none"	
		    }


#open first page of your pdf file in a loop
foreach ($pdf in $pdfs){
$reader = New-Object iTextSharp.text.pdf.pdfreader($($path+"\"+$pdf))
$text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, 1)
$reader.Close()

#search for lastname from xlsx content in a pdf file
$match = [regex]::matches($text, $lastname)
if ($match.Success) #if success - check if the firstname matches
{	
	$matchFirstName = [regex]::matches($text, $firstname)
	  if($matchFirstName.Success)
      {
        if ($firstname.length -eq 0) {
        $filename = $city + "_" + $toFilename + ".pdf"
        }
        else {
	            $filename = $city + "_" + $toFilename + "_" + $firstname + ".pdf"
            }
            
  #replace diacritical marks
	$filename = $filename -replace( "Ą","A" )`
                        -replace( "Ć","C" )`
                        -replace( "Ę","E" )`
                        -replace( "Ł","L" )`
                        -replace( "Ń","N" )`
                        -replace( "Ó","O" )`
                        -replace( "Ś","S" )`
                        -replace( "Ż","Z" )`
                        -replace( "Ź","Z" )`
                        -replace( ",","." )`
                        -replace( " ","." )`
	                      -replace("\.\.",".")

#copy pdf file to directory and change it's name
copy-item -path ($($path+"\"+$pdf)) -destination $($path+"\out\" + $filename)
#move source files to another directory
move-item -path ($($path+"\"+$pdf)) -destination $($path+"\done\" + $($pdf.Name) -force

    $pdfs = $pdfs | Where-Object {$_ -ne $pdf}
}
}
}
}
