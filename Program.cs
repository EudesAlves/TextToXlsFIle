
TransferTxtToExcelData();

void TransferTxtToExcelData()
{
    Console.WriteLine("Starting ...");

    TextToExcelDataManager txtToExelManager = new TextToExcelDataManager();
    txtToExelManager.WriteTextFileDataToExcel();
    
    /*
    TxtDataManager txtManager = new TxtDataManager();
    txtManager.GetDataFromTxtFiles();
    txtManager.FillLinesOccurrencesCountInTextFiles();
    txtManager.ShowLinesOccurrencesCountInTextFiles();

    XlsDataManagerOffice xlsManager = new XlsDataManagerOffice();
    string sheetName = "icemakerzero";
    int row = 25;
    int column = 8;
    xlsManager.GetCellsFromColumn(sheetName, row, column);

    xlsManager.WriteCellData( sheetName, 27, "H", "");
    xlsManager.WriteCellData( sheetName, 28, "H", "");
    xlsManager.WriteCellData( sheetName, 29, "H", "");
    xlsManager.WriteCellData( sheetName, 30, "H", "");
    xlsManager.WriteCellData( sheetName, 31, "H", "");
    xlsManager.WriteCellData( sheetName, 32, "H", "");
    */
    Console.ReadLine();
}


