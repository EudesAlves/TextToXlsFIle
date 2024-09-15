using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Data.OleDb;

public class XlsDataManagerOleDb
{
    static string baseDirectory = AppContext.BaseDirectory;
    static string xlsFileName = "Itens da live.xlsx";
    private string xlsFilePath = baseDirectory + xlsFileName;

    public XlsDataManagerOleDb()
    {}

    public XlsDataManagerOleDb(string _xlsFileName)
    {
        xlsFileName = _xlsFileName;
        xlsFilePath = baseDirectory + xlsFileName;
    }

    public void ShowFilePath()
    {
        Console.WriteLine("" + xlsFilePath);
    }

    public DataTable ReadXlsData()
    {
        DataTable dtResult = new DataTable();
        string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={xlsFilePath};Extended Properties=\'Excel 12.0 Xml;HDR=YES\'";

        using (var oleDbConn = new OleDbConnection(connectionString))
        {
            try 
            {
                oleDbConn.Open();
                string tableName = "Icemakerzero";
                string query = $"SELECT * FROM [{tableName}$]";

                using (var oleDbComm = new OleDbCommand())
                {
                    oleDbComm.CommandText = query;
                    oleDbComm.Connection = oleDbConn;

                    using (var oleDbAdap = new OleDbDataAdapter())
                    {
                        oleDbAdap.SelectCommand = oleDbComm;
                        oleDbAdap.Fill(dtResult);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{ex.Message}");
                throw ex;
            }
            finally
            {
                oleDbConn.Close();
                oleDbConn.Dispose();
            }

            return dtResult;
        }
    }

    public void ShowXlsData()
    {
        var xlsDataTable = new DataTable();
        xlsDataTable = ReadXlsData();

        DataRow row = xlsDataTable.Rows[24];
        var value = row.ItemArray[1];
        Console.WriteLine(string.Format("", value.ToString()));


        
        Console.WriteLine($"\n Table name: {xlsDataTable.TableName}");

        foreach (DataColumn column in xlsDataTable.Columns)
        {
            Console.WriteLine(string.Format("{0} ", column.ColumnName));
        }

        foreach (DataRow rows in xlsDataTable.Rows)
        {
            //var record1 = rows
            //Console.WriteLine(string.Format("", record1));
            
            
            Console.WriteLine("_________ new Record __________");
            foreach (var record in rows.ItemArray)
            {
               Console.WriteLine($"{record}");
            }
        }
    }
}