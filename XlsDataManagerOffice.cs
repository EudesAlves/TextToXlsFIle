using System;
using System.Data;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

public class XlsDataManagerOffice
{
    static string baseDirectory = AppContext.BaseDirectory;
    static string xlsFileName = "Itens da live.xlsx";
    private string xlsFilePath = baseDirectory + xlsFileName;
    public Dictionary<string, int> itemsFromCells = new Dictionary<string, int>();
    public Dictionary<string, string> keyWordsToTranslateItems = new Dictionary<string, string>()
    {
        { "ferro", "barra de ferro" },
        { "bronze", "barra de bronze" },
        { "prata", "barra de prata" },
        { "ouro", "barra de ouro" },
        { "platina", "barra de platina" },
        { "gema", "gema misteriosa" },
        { "misteriosa", "gema misteriosa" },
        { "som", "som na live" },
        { "sons", "som na live" },
        { "musica", "musica na live" },
        { "musicas", "musica na live" },
        { "evento", "evento na live" },
        { "eventos", "evento na live" },
        { "desejo", "desejo" },
        { "desejos", "desejo" },
        { "ticket", "ticket exta" },
        { "rops", "vale sorteio de rops" },
        { "jackpot", "jackpot no sorteio na roleta" },
        { "level", "level roleplay na roleta" },
        { "sorteio", "sorteio de jogo" }
    };

    public void FillItemsDictionary(string sheetName, int row, int column)
    {
        itemsFromCells = new Dictionary<string, int>();
        List<string> cellsData = new List<string>();
        Dictionary<string, int> dictItem = new Dictionary<string,int>();

        cellsData = GetCellsFromColumn(sheetName, row, column);

        foreach(string cell in cellsData)
        {
            dictItem = SeparateNumberFromText(cell);
            if(!itemsFromCells.Keys.Contains(dictItem.Keys.First<string>()))
            {
                string translatedText = TranslateXlsItems(dictItem.Keys.First<string>().ToLower());

                itemsFromCells.Add(dictItem.Keys.First<string>(), dictItem.Values.First<int>());
            }
        }
    }
    
    public List<string> GetCellsFromColumn(string sheetName, int row, int column)
    {
        List<string> cellsData = new List<string>();
        string cellsValue = GetCellData(sheetName, row, column);

        while(!string.IsNullOrEmpty(cellsValue))
        {
            cellsData.Add(cellsValue);
            row ++;
            cellsValue = GetCellData(sheetName, row, column);
        }

        Console.WriteLine($"Linhas obtidas: {sheetName}  {cellsData.Count}");
        return cellsData;
    }

    public string GetCellData(string sheetName, int row, int column)
    {
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workBook;
        Excel.Worksheet workSheet;

        workBook = excelApp.Workbooks.Open(xlsFilePath);
        workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];

        string cellData = Convert.ToString(workSheet.Cells[row, column].Value);
        workBook.Close();
        excelApp.Quit();
        
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workBook);

        KillExcelProcess();

        Console.WriteLine("Cell " + cellData);
        return cellData;
    }

    public void WriteCellsOnColumn(string sheetName, int row, string column, List<string> data)
    {
        /*
        foreach(string cell in data)
        {
            WriteCellData(sheetName, row, column, cell);
            row ++;
        }*/
        string[] arrayData = data.ToArray();
        WriteRangeCellsOnColumn(sheetName, row, column, arrayData);
    }

    public void WriteRangeCellsOnColumn(string sheetName, int initialRow, string column, string[] data)
    {
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workBook;
        Excel.Worksheet workSheet;

        workBook = excelApp.Workbooks.Open(xlsFilePath);
        workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];

        int row = initialRow;
        string rangeString = "";

        foreach (string text in data)
        {
            rangeString = column + row.ToString() + ":" + column + row.ToString();
            Excel.Range cellRange = workSheet.Range[rangeString];
            cellRange.Value = text;

            row ++;
        }

        excelApp.DisplayAlerts = false;
        
        workBook.SaveAs(xlsFilePath);
        workBook.Close();
        excelApp.Quit();
        
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workBook);

        KillExcelProcess();
    }
    
    public void WriteCellData(string sheetName, int row, string column, string data)
    {
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workBook;
        Excel.Worksheet workSheet;

        workBook = excelApp.Workbooks.Open(xlsFilePath);
        workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];

        string rangeString = column + row.ToString() + ":" + column + row.ToString();
        Excel.Range cellRange = workSheet.Range[rangeString];
        cellRange.Value = data;

        excelApp.DisplayAlerts = false;
        
        workBook.SaveAs(xlsFilePath);
        workBook.Close();
        excelApp.Quit();
        
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workBook);

        KillExcelProcess();
    }

    public Dictionary<string, int> SeparateNumberFromText(string data)
    {
        Dictionary<string, int> numberText = new Dictionary<string, int>();
        int number = 1;
        string text = data;
        int spaceIndex = data.IndexOf(' ');
        string firstSlice = "1";
        if(spaceIndex > 0)
        {
            firstSlice = data.Substring(0, spaceIndex);
        }

        try{
            number = Convert.ToInt32(firstSlice);
            text = data.Substring(spaceIndex +1, data.Length -(spaceIndex +1));
        }
        catch(Exception ex)
        {
            number = 1;
            text = data;
            Console.WriteLine(ex.Message);
        }
        finally
        {
            
            numberText.Add(text, number);
        }

        Console.WriteLine($"number:{number.ToString()} _ text:{text}");

        return numberText;
    }

    public void KillExcelProcess()
    {
        try
        {
            Process[] excelProcesses = Process.GetProcessesByName("excel");
        foreach (Process p in excelProcesses)
        {
            if (string.IsNullOrEmpty(p.MainWindowTitle)) // use MainWindowTitle to distinguish this excel process with other excel processes 
            {
                p.Kill();
            }
        }
        }
        catch(Exception ex)
        {
            Console.WriteLine(ex.Message);
        }        
    }

    public string TranslateXlsItems(string originText)
    {
        foreach(string itemKey in keyWordsToTranslateItems.Keys)
        {
            if(originText.Contains(itemKey))
            {
                return keyWordsToTranslateItems[itemKey];
            }
        }

        return originText;
    }
}