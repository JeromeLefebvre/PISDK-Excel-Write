# PISDK-Excel-Write
A short sample Excell workbook showing how to write to a PI Server using PI-SDK and Visual Basic within Microsoft Excel.

## Requirement
The script has been tested with Microsoft Office 2016 and with PI SDK 2014 R2.

## Getting Started
To use the script, simply open the PISDKExcelSample_Write.xlsm enter a tag, timestamp and a value. Clicking PutValue will then send that value to the PI Data Archive.
If you want to send the same data to multiple PI Servers, delimit the PI Servers by a pipe character.

## The code
You can view the code in either Module1.bas (the main code) or ThisWorkbook.cls (which contains a small section to automate sending the value). The code comments are written in both English and Japanese.