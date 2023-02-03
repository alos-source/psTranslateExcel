# psTranslateExcel
psTranslateExcel is a PowerShell script that automatically performs the translation of text from an Excel spreadsheet using the DeepL API.

## Requirements
- Microsoft Excel
- PowerShell 5.0 or higher
- Internet connection

## Usage
The script can be executed through PowerShell. It uses the following parameters:
- ExcelFilePath (required) - the path to the Excel file to be translated.
- SheetIndex (required) - the index of the worksheet in the Excel file to be translated.
- SourceColumn (required) - the column in the Excel sheet with the source texts.
- TargetColumn (required) - The column in the Excel sheet where the translations should be written.
- APIKey (required) - The API key for the DeepL API, see also https://www.deepl.com/de/docs-api/

The script translates the texts in the SourceColumn and writes the translations to the TargetColumn. The changes are saved in the Excel file.

## Example
The following example shows how to call the script with the appropriate parameters:
        TranslateExcelColumn -ExcelFilePath "C:\myfile.xlsx" -SheetIndex 1 -SourceColumn "A" -TargetColumn "B" -APIKey "YOUR_API_KEY"
