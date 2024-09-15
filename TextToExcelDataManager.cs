using System;
using System.Data;

public class TextToExcelDataManager
{
    public void WriteTextFileDataToExcel()
    {
        Dictionary<string, Dictionary<string, int>> txtLinesData = new Dictionary<string, Dictionary<string, int>>();
        Dictionary<string, int> xlsLinesData = new Dictionary<string, int>();
        TxtDataManager txtManager = new TxtDataManager();
        XlsDataManagerOffice xlsManager = new XlsDataManagerOffice();
        int initialRow = 25;
        int columnNumber = 8;
        string columnLetter = "H";

        txtManager.GetDataFromTxtFiles();
        txtManager.FillLinesOccurrencesCountInTextFiles();
        txtLinesData = txtManager.linesFilesOccurrencesCount;
        foreach(string fileName in txtLinesData.Keys)
        {
            int extensionIndex = fileName.IndexOf(".");
            string sheetName = fileName.Substring(0, extensionIndex);
            xlsManager.FillItemsDictionary(sheetName, initialRow, columnNumber);
            xlsLinesData = xlsManager.itemsFromCells;

            Dictionary<string, int> mergedData = MergeTwoDictsOfItems(txtLinesData[fileName], xlsLinesData);
            List<string> resultListData = new List<string>();
            foreach(string text in mergedData.Keys)
            {
                resultListData.Add(mergedData[text].ToString() + " " + text);
            }

            xlsManager.WriteCellsOnColumn(sheetName, initialRow, columnLetter, resultListData);
        }

    }

    Dictionary<string, int> MergeTwoDictsOfItems(Dictionary<string, int> originData, Dictionary<string, int> destinyData)
    {
        Dictionary<string, int> resultDict = new Dictionary<string, int>();

        foreach(string text in destinyData.Keys)
        {
            if(!originData.Keys.Contains(text))
            {
                resultDict.Add(text, destinyData[text]);
            }
        }
        
        foreach(string text in originData.Keys)
        {
            if(destinyData.Keys.Contains(text))
            {
                int total = originData[text] + destinyData[text];
                resultDict.Add(text, total);
            }
            else
            {
                resultDict.Add(text, originData[text]);
            }
        }

        return resultDict;
    }    
}