
TransferTxtToExcelData();

void TransferTxtToExcelData()
{
    Console.WriteLine("Starting ...");

    TextToExcelDataManager txtToExelManager = new TextToExcelDataManager();
    txtToExelManager.WriteTextFileDataToExcel();
    
    Console.ReadLine();
}


