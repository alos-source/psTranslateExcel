# psTranslateExcel
psTranslateExcel is a PowerShell script that automatically performs the translation of text from an Excel spreadsheet using the DeepL API.

## Requirements
- Microsoft Excel
- PowerShell 5.0 or higher
- Internet connection
- DeepL-Account to get an API-Key
- API-Key is Configured as an User-Environment-Variable in the Operating System

## Usage
The script can be executed through PowerShell. It uses the following parameters:
- ExcelFilePath (required) - the path to the Excel file to be translated.
- SheetIndex (required) - the index of the worksheet in the Excel file to be translated.
- SourceColumn (required) - the column in the Excel sheet with the source texts.
- TargetColumn (required) - The column in the Excel sheet where the translations should be written.
- APIKey (required) - The API key for the DeepL API, see also https://www.deepl.com/de/docs-api/

The script translates the texts in the SourceColumn and writes the translations to the TargetColumn. The changes are saved in the Excel file. The script can be called in powershell using “dot sourcing”:
```
    . ./translateExcel.ps1
```
## Example
The following example shows how to call the script with the appropriate parameters:
```
    TranslateExcelColumn -ExcelFilePath "C:\myfile.xlsx" -SheetIndex 1 -SourceColumn "A" -TargetColumn "B" -APIKey "YOUR_API_KEY"
```

## Note
Please note that the script uses the DeepL API and therefore may transfer data over the Internet. Make sure that you are aware of what data is being transferred and that you agree with the privacy policy of the DeepL API.
