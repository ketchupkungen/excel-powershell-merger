#Keep in mind, it´s just a Prototype.
#First I create a workbook named DONE, with a date named sheet
$xl = New-Object -ComObject Excel.Application
$File1 = 'C:\Users\evil_\Documents\excel-powershell-merger\Result.xlsm'  #File we want to Copy from
$File2 = 'C:\Users\evil_\Documents\excel-powershell-merger\Prototyping.xlsx'  #File we want to Copy from

Add-Type -AssemblyName System.Windows.Forms
$msgBoxInput=[System.Windows.Forms.MessageBox]::Show("Would you like to create a merge between`n
    $File1`n and`n $File2`n into a new file named`n $(get-date -f yyyy-MM-dd).xlsx"`
    ,"File merger",[System.Windows.Forms.MessageBoxButtons]::YesNo)
    switch  ($msgBoxInput) {

      'Yes' {
      ## Proceed
      }

      'No' {
      ## End script
      $xl.quit()
      exit
      }
}
$xl.displayAlerts = $false
$xl.Visible = $false
$wb = $xl.Workbooks.add()
#$workbook.Worksheets.Item(1).Delete() #Excel worksheets sometimes come with too many worksheets
$ws1 = $wb.worksheets.Item(1)
$ws1.name = $(get-date -f yyyy-MM-dd) #"Sheet1"
$wb.SaveAs("C:\Users\evil_\Documents\excel-powershell-merger\$(get-date -f yyyy-MM-dd).xlsx")

#Then I copy from the old file and paste to the new
#UsedRange would be nice to implement, instead of selecting col/row:col/row
#And figure out how to paste below active cells
$range1="A1:P46"
$range2="A1:A10" #Temporary
$range3="A1"
$File3 = "C:\Users\evil_\Documents\excel-powershell-merger\$(get-date -f yyyy-MM-dd).xlsx" #File we want to Paste to
$wb1 = $xl.workbooks.open($File1, $null, $true)
$wb2 = $xl.workbooks.open($File2, $null, $true)
$wb3 = $xl.workbooks.open($File3)

#File1-CP
$ws1 = $wb1.WorkSheets.item(1)
$ws1.activate()
$range = $ws1.Range($range1).Copy()

$ws3 = $wb3.Worksheets.item(1)
$ws3.activate()
$x=$ws3.Range($range3).Select()

$ws3.Paste()
#/File1-CP

#File2-CP
$ws2 = $wb2.WorkSheets.item(1)
$ws2.activate()
$range = $ws2.Range($range2).Copy()

$ws3 = $wb3.Worksheets.item(1)
$ws3.activate()
$x=$ws3.Range($range3).Select()

$ws3.Paste()
$wb3.Save()
#/File2-CP

$wb1.close($false)
$wb2.close($true)
$xl.quit()
#spps -n excel

#Message that it was  successful
#Add-Type -AssemblyName System.Windows.Forms
$oReturn=[System.Windows.Forms.MessageBox]::Show("Finnished!",`
    "File merger",[System.Windows.Forms.MessageBoxButtons]::OK)
