using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection.Metadata.Ecma335;

public class TxtDataManager
{
    static string baseDirectory = AppContext.BaseDirectory;
    private string directoryPath = Path.Combine(baseDirectory, "TxtFiles");
    private Dictionary<string, List<string>> filesContent = new Dictionary<string, List<string>>();

    public Dictionary<string, Dictionary<string, int>> linesFilesOccurrencesCount = new Dictionary<string, Dictionary<string, int>>();

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
        { "desejo", "desejo" },
        { "ticket", "ticket exta" },
        { "rops", "vale sorteio de rops" },
        { "jackpot", "jackpot no sorteio na roleta" },
        { "level", "level roleplay na roleta" },
        { "sorteio", "sorteio de jogo" }
    };

    public void GetDataFromTxtFiles()
    {
        if(!Directory.Exists(directoryPath))
        {
            Console.WriteLine($"Directory not Found: {directoryPath}");
            return;
        }
        
        Console.WriteLine($"Reading directory {directoryPath} ...");

        string[] txtFiles = Directory.GetFiles(directoryPath, "*.txt");
        foreach (string file in txtFiles)
        {
            string fileName = Path.GetFileName(file);
            string[] lines = File.ReadAllLines(file);

            List<string> reviewdLines = GetValidTextFromLines(new List<string>(lines));
            
            filesContent.Add(fileName, reviewdLines);
        }
    }

    List<string> GetValidTextFromLines(List<string> list)
    {
        List<string> reviewdLines = new List<string>();
        string prefixToRemove = "#";
        string sufixToRemove = " nwg";
        string invalidLineText = "NADA";
        int prefixIndex = 0;
        int sufixIndex = 0;
        int invalidTextIndex = -1;

        foreach(string line in list)
        {
            invalidTextIndex = line.IndexOf(invalidLineText);
            if(invalidTextIndex == -1)
            {
                string lineLower = line.ToLower();
                prefixIndex = lineLower.IndexOf(prefixToRemove) > -1 ? lineLower.IndexOf(prefixToRemove)+3 : 0;
                sufixIndex = lineLower.IndexOf(sufixToRemove) > -1? lineLower.IndexOf(sufixToRemove) : 0;
                int finalIndex = sufixIndex > prefixIndex ? sufixIndex : lineLower.Length;

                string validText = lineLower.Substring(prefixIndex, finalIndex - prefixIndex);
                string translatedText = TranslateTxtItems(validText);

                reviewdLines.Add(translatedText);
            }
        }

        return reviewdLines;
    }

    List<string> RemovePrefixFromList(List<string> list, char lastPrefixChar)
    {
        List<string> reviewdLines = new List<string>();
        foreach(string line in list)
        {
            string slicedText = line;
            int indexChar = line.IndexOf(lastPrefixChar);
            if(indexChar >= 0)
            {
                slicedText = line.Substring(indexChar+1, line.Length - (indexChar+1));
            }
            reviewdLines.Add(slicedText);
        }

        return reviewdLines;
    }

    public void ShowFilesList()
    {
        Console.WriteLine($"Showing files in directory {directoryPath} ...");
        foreach (string fileName in filesContent.Keys)
        {
            Console.WriteLine("" + fileName);
        }
    }

    public void FillLinesOccurrencesCountInTextFiles()
    {
        foreach (string fileName in filesContent.Keys)
        {
            List<string> lines = filesContent[fileName];
            linesFilesOccurrencesCount.Add(fileName, SumLinesOccurrencesInList(lines));
        }
    }

    public Dictionary<string, int> SumLinesOccurrencesInList(List<string> items)
    {
        Dictionary<string, int> countOccurences = new Dictionary<string, int>();

        foreach (string item in items)
        {
            if(countOccurences.ContainsKey(item))
            {
                countOccurences[item] ++;
            }
            else { 
                countOccurences.Add(item, 1);
            }
        }

        return countOccurences;
    }

    public void ShowLinesOccurrencesCountInTextFiles()
    {
        Console.WriteLine($"Showing Lines Ocurrences count: /n");
        foreach (string fileName in linesFilesOccurrencesCount.Keys)
        {
            Console.WriteLine("________ " + fileName + " _______ ");

            Dictionary<string , int> linesOccurencesCount = linesFilesOccurrencesCount[fileName];
            foreach (string line in linesOccurencesCount.Keys)
            {
                Console.WriteLine(line + " - " + linesOccurencesCount[line]);
            }
        }
    }

    public string TranslateTxtItems(string originText)
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