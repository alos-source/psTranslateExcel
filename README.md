# psTranslateExcel

## How-To Call
The function has the parameters:
        - ExcelFilePath [string](Mandatory=$true)
        - SheetIndex [int](Mandatory=$true)
        - SourceColumn [string](Mandatory=$true)]
        - TargetColumn [string](Mandatory=$true)]
        - $APIKey [string](Mandatory=$true)]
        
An example call could look like this:
' TranslateExcelColumn -ExcelFilePath "D:\UserData\myUser\translate.xlsx" -SheetIndex 2 -SourceColumn "D" -TargetColumn "E" -APIKey $APIKEY
