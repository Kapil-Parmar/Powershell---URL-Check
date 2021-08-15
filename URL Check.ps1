#get input sheet path
$inputSheetPath =  $PSScriptRoot +'\InputSheet\InputSheet.xlsm'

#generate report sheet name
$currentDate = Get-Date -Format "ddMMyyyy_HHmmss"
$reportSheetPath =  $PSScriptRoot +'\ReportSheet\ReportSheet'+ $currentDate +'.xlsm'

#read excel
$xl = New-Object -ComObject "Excel.Application"
$wb = $xl.Workbooks.Open($inputSheetPath)
$ws = $wb.Sheets.Item(1)

#for each row in input sheet
for($i = 2; $i -le ($ws.UsedRange.Rows).count; $i++)
{
  try 
  {
    #get url from input sheet
    $url = $ws.cells.item($i,1).text

    #make http request
    $req = [system.Net.WebRequest]::Create($url)

    #get response from http request
    $res = $req.GetResponse()

    #get status code from response
    $StatusCode = [int]$res.StatusCode

    #update status code in excel
    $ws.cells.item($i,2) = $StatusCode
    Start-Sleep -Seconds 2
  } 
  catch 
  {
    #update exception message in excel incase of exception
    #$res = $PsItem.Exception.Response
    #$errorMessage = $Error[0].Exception.GetType().FullName
    $errorMessage = $PSItem.Exception.Message
    $ws.cells.item($i,2) = $errorMessage
  }
}

#execute macro which will update cell color of status code which are not 200 as red
$xl.Run("FormatCells")

#save report sheet
[void]$wb.SaveAs($reportSheetPath)

#release excel com object
[void]$wb.Close()
[void]$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null