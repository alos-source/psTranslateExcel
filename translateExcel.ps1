function TranslateExcelColumn {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string] $ExcelFilePath,

        [Parameter(Mandatory=$true)]
        [int] $SheetIndex,

        [Parameter(Mandatory=$true)]
        [string] $SourceColumn,

        [Parameter(Mandatory=$true)]
        [string] $TargetColumn,

        [Parameter(Mandatory=$true)]
        [string] $APIKey
    )

    # Load the required assembly for working with Excel files
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.Excel')

    # Create an instance of Excel
    $excel = New-Object -ComObject Excel.Application

    # Open the Excel file
    $workbook = $excel.Workbooks.Open($ExcelFilePath)

    # Select the specified worksheet
    $worksheet = $workbook.Sheets.Item($SheetIndex)

    # Get the data from the source column of the worksheet with the source texts, skip the first line as header
    
     for ($row = 2; $row -le $worksheet.UsedRange.Rows.Count; $row++) {
        $cellValue = $worksheet.Cells.Item($row, ${SourceColumn}).Value2
        Write-Output $cellValue
               # Use the DeepL API to translate the data
        $url = "https://api-free.deepl.com/v2/translate?auth_key=$APIKey&target_lang=EN&text=" + [System.Uri]::EscapeDataString($cellValue)
        $result = Invoke-RestMethod -Uri $url
        # Write-Output $result
        $translatedText = $result.translations[0].text
        Write-Output $translatedText
        
        # Write the translation result to the target column
        $worksheet.Cells.Item($row, $TargetColumn) = $translatedText.ToString()
 
     }

    # Save the changes and close the Excel file
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
