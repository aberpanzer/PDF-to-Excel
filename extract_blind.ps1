function convert-PDFtoText {
	param(
		[Parameter(Mandatory=$true)][string]$file
	)	
	Add-Type -Path "D:\bupotpdftoxls\itextsharp.dll"
	$pdf = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $file
	for ($page = 1; $page -le $pdf.NumberOfPages; $page++){
		$text=[iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($pdf,$page)
		Write-Output $text
	}	
	$pdf.Close()
}



$excel = New-Object -Com Excel.Application
$Excel.Visible = $True 

$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Add()
$ws.name = "Bupot"

$ws.Cells.Item(1,1).Value="No."
$ws.Cells.Item(1,2).Value="NPWP"
$ws.Cells.Item(1,3).Value="NPWP 16"
$ws.Cells.Item(1,4).Value="NITKU"
$ws.Cells.Item(1,5).Value="NAMA Customer"
$ws.Cells.Item(1,6).Value="Nomor Bukti Potong"
$ws.Cells.Item(1,7).Value="Nomor Dokumen"
$ws.Cells.Item(1,8).Value="Tanggal Bukti Potong"
$ws.Cells.Item(1,9).Value="Mapping"
$ws.Cells.Item(1,10).Value="Jenis Penghasilan Atas"
$ws.Cells.Item(1,11).Value="DPP"
$ws.Cells.Item(1,12).Value="Tarif"
$ws.Cells.Item(1,13).Value="PPh"

$irow=1

$files = Get-ChildItem -Path "D:\bupotpdftoxls\File" -Filter "*.pdf"
foreach ($file in $files) {
    # Perform actions on each .txt file
    $text = convert-PDFtoText $file.FullName
		
		
	$irow=$irow+1
	
	#No	
	$norow=$irow-1
	$ws.Cells.Item($irow,1).Value=[string]$norow
	
	#NPWP
	$NPWP1=$text.IndexOf("C. IDENTITAS PEMOTONG/PEMUNGUT")
	$NPWP2=$text.IndexOf("C.1 NPWP")
	$NPWP=$text.substring($NPWP1 + 32,$NPWP2 - ($NPWP1 + 32) - 1)
	$ws.Cells.Item($irow,2).Value="'"+$NPWP.substring(1,15)
	
	#NPWP 16	
	$ws.Cells.Item($irow,3).Value="'"+$NPWP.substring($NPWP.Length - 16, 16)
	
	#NITKU
	$NITKU1=$text.IndexOf("C.2 NITKU")
	$ws.Cells.Item($irow,4).Value="'"+$text.substring($NITKU1 + 10, 22)
	
	#Nama Customer
	$NAMA1=$text.IndexOf("C.3 Nama Wajib Pajak")
	$NAMA2=$text.IndexOf("C.4 Tanggal")
	$ws.Cells.Item($irow,5).Value=$text.substring($NAMA1 + 21,$NAMA2 - ($NAMA1 + 21) - 1)
	
	#Nomor Bukti Potong
	$NOBUPOT1=$text.IndexOf("NOMOR      : ")
	$NOBUPOT2=$text.IndexOf("PPh Final")
	$NOBUPOT=$text.substring($NOBUPOT1 + 13,$NOBUPOT2 - ($NOBUPOT1 + 13) - 1)
	$ws.Cells.Item($irow,6).Value="'"+$NOBUPOT.Replace(" ","")
	
	#Nomor Dokumen 
	$NODOK1=$text.IndexOf("Dokumen Referensi  : Nomor Dokumen ")
	$NODOK2=$text.IndexOf("Nama Dokumen ")
	$NODOK3=($NODOK2 - 33) - ($NODOK1 + 35)
	$NODOK4=$text.substring($NODOK1+35, $NODOK3)
	
	$NODOK5=$NODOK4.IndexOf("`n")
	if ($NODOK5 -ge 0) {
	 $NODOK6=$NODOK4.substring(1, $NODOK5-1)
	 $ws.Cells.Item($irow,7).Value="'"+$NODOK6.Replace(" ","")	
	}else{
	 $ws.Cells.Item($irow,7).Value="'"+$NODOK4.Replace(" ","")		
	}
	
	#Tanggal Bukti Potong
	$TGLBUPOT1=$text.IndexOf("C.4 Tanggal")
	$TGLBUPOT2=$text.substring($TGLBUPOT1 + 14,15)
	$TGLBUPOT="'"+$TGLBUPOT2.Replace(" ","")
	$ws.Cells.Item($irow,8).Value=$TGLBUPOT.substring(3,2)+"/"+$TGLBUPOT.substring(1,2)+"/"+$TGLBUPOT.substring(5,4)
		
	#Mapping
	$NILAI1=$text.IndexOf("B.1 B.2 B.4 B.5 B.6")
	$NILAI2=$text.IndexOf("Keterangan Kode Objek Pajak")
	$NILAI=$text.substring($NILAI1 + 21,$NILAI2 - ($NILAI1 + 21) - 1)
	
	$ws.Cells.Item($irow,9).Value="'"+$text.substring($NILAI1 + 21, 9)
			
	#Jenis Penghasilan Atas
	$JPA1=$text.IndexOf("Keterangan Kode Objek Pajak")
	$JPA2=$text.IndexOf("B.7")	
	$ws.Cells.Item($irow,10).Value=$text.substring($JPA1 + 32,$JPA2 - ($JPA1 + 32) - 1)
		
	
	#PPh
	$PPh1=$text.substring($NILAI1 + 31,$NILAI2 - ($NILAI1 + 31) - 1)
	$PPh2=$PPh1.IndexOf("-")
	$PPh3=$PPh1.substring(0,$PPh2 - 2)
	$ws.Cells.Item($irow,13).Value=$PPh3
		
	#DPP
	$DPP1=$PPh1.IndexOf("-")
	$DPP2=$PPh1.Length
	$DPP3=$PPh1.substring($DPP1+6,($DPP2)-($DPP1+6))	
	$DPP4=$DPP3.IndexOf(",00 ")
	$DPP5=$DPP3.substring(0,$DPP4+3)
	$ws.Cells.Item($irow,11).Value=$DPP5	
	
	#Tarif	
	$TARIF1=$DPP3.Length-($DPP4+4)
	$TARIF=$DPP3.Substring($DPP4+4, $TARIF1)	
	$ws.Cells.Item($irow,12).Value=$TARIF.Replace(" ","")
}

Write-Host “------------------------------”
Write-Host “Proses Konversi sudah Selesai.”
Write-Host “------------------------------”

# Senang berbagi, Jika ada kesulitan info saja berlinasep@gmail.com