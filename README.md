# Text To Xls File

Dotnet Console Program
This program:
Reads the lines of all .txt files in a directory;
Adds the same lines and saves them in a dictionary containing the item and quantity;
For each txt file, reads the lines from column "H" in the Sheet
corresponding to the name of the txt file;
Merges the lines from the spreadsheet with the lines from the txt;
Saves the resulting items in the xls file, in column "H"
of the Sheet corresponding to the name of the txt file;

The class XlsDataManagerOffice is used to manage the xls file;
The class XlsDataManagerOleDb it's not being used;

XLS destiny file: "./Itens da live.xlsx"
Txt source files directory: "./TxtFiles"


## Dependencies

Dotnet Runtime
   https://dotnet.microsoft.com/pt-br/download/dotnet/thank-you/runtime-8.0.7-windows-x64-installer

XlsDataManagerOffice
    Microsoft.Office.Interop.Excel
    https://www.microsoft.com/en-us/microsoft-365/download-office
    Microsoft Office 2010: Primary Interop Assemblies Redistributable
    https://www.microsoft.com/en-us/download/details.aspx?id=3508

XlsDataManagerOleDb
    System.Data.OleDb
    Microsoft Access Database Engine 2016 Redistributable
    https://www.microsoft.com/en-us/download/details.aspx?id=54920


### Run the Program
\bin\Release\net8.0\publish\TextToXlsFIle.exe


#### TODO
- The txt file directory needs to be parameterized;
- The xls file directory needs to be parameterized;
- Column "H" of the xls spreadsheet needs to be parameterized; 

